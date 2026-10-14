"""Checks that no workbook under a folder carries personal data in its metadata or parts.

    python -I verify-metadata.py <folder> [--forbid TEXT]...

For every .xlsx: docProps/core.xml creator and lastModifiedBy and docProps/app.xml Company must
be empty (what Excel's "remove personal information on save" writes) or the neutral value, there
must be no docProps/custom.xml, and no part may contain a user folder (absPath or any other) or
one of the forbidden strings.
For every .xls: the raw bytes (Latin-1 and UTF-16LE) may contain none of the forbidden strings.

Forbidden strings are this machine's user name, computer name, domain and profile folder (read
from the environment, so nothing personal is written into this script) plus every --forbid.
Exits 1 and lists the problems if anything is found. Standard library only.
"""

import os
import re
import sys
import zipfile

NEUTRAL = "GridCraft contributors"


def forbidden_terms(extra):
    terms = set(t for t in extra if t)
    for var in ("USERNAME", "COMPUTERNAME", "USERDOMAIN", "USER", "HOSTNAME", "LOGNAME"):
        v = os.environ.get(var, "").strip()
        if len(v) >= 3:
            terms.add(v)
    profile = os.environ.get("USERPROFILE") or os.environ.get("HOME") or ""
    if profile:
        terms.add(os.path.basename(profile.rstrip("\\/")))
    # A display name like "First Last" is also checked word by word.
    for t in list(terms):
        for w in re.split(r"\s+", t):
            if len(w) >= 4:
                terms.add(w)
    return sorted(terms, key=str.lower)


def tag(xml, name):
    m = re.search(r"<%s>(.*?)</%s>" % (name, name), xml, re.S)
    return m.group(1) if m else None


def check_xlsx(path, terms):
    problems = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        if "docProps/custom.xml" in names:
            problems.append("has docProps/custom.xml")
        core = z.read("docProps/core.xml").decode("utf-8") if "docProps/core.xml" in names else ""
        app = z.read("docProps/app.xml").decode("utf-8") if "docProps/app.xml" in names else ""
        for name, xml in (("dc:creator", core), ("cp:lastModifiedBy", core), ("Company", app)):
            v = tag(xml, name)
            if v is not None and v not in ("", NEUTRAL):
                problems.append("%s is %r" % (name, v))
        for n in names:
            text = z.read(n).decode("utf-8", "replace")
            if re.search(r"[A-Za-z]:\\Users\\|/Users/|/home/", text, re.I):
                problems.append("%s contains a user folder path" % n)
            low = text.lower()
            for t in terms:
                if t.lower() in low:
                    problems.append("%s contains a forbidden string" % n)
                    break
    return problems


def check_xls(path, terms):
    data = open(path, "rb").read().lower()
    problems = []
    for t in terms:
        for enc in ("latin-1", "utf-16-le"):
            try:
                needle = t.lower().encode(enc)
            except UnicodeEncodeError:
                continue
            if needle in data:
                problems.append("contains a forbidden string (%s)" % enc)
    if re.search(rb"[a-z]:\\users\\", data) or re.search(rb"[a-z]\x00:\x00\\\x00u\x00s\x00e\x00r\x00s\x00", data):
        problems.append("contains a user folder path")
    return problems


def main(argv):
    args = argv[1:]
    extra = []
    folder = None
    while args:
        a = args.pop(0)
        if a == "--forbid":
            extra.append(args.pop(0))
        else:
            folder = a
    if not folder:
        print(__doc__)
        return 2
    terms = forbidden_terms(extra)
    bad = 0
    count = 0
    for dirpath, _, files in os.walk(folder):
        for f in sorted(files):
            p = os.path.join(dirpath, f)
            ext = f.lower().rsplit(".", 1)[-1]
            if ext == "xlsx":
                problems = check_xlsx(p, terms)
            elif ext == "xls":
                problems = check_xls(p, terms)
            else:
                continue
            count += 1
            for pr in problems:
                bad += 1
                print("METADATA  %s: %s" % (os.path.relpath(p, folder), pr))
    print("metadata: %d files checked, %d problems" % (count, bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
