# AGENTS.md: working on the GridCraft corpus

## 1. Read first

1. [`README.md`](README.md): what this repository is, the layout, checking, regenerating,
   metadata rules, licence.
2. GridCraft's `AGENTS.md` (https://github.com/storytold/gridcraft) and its `check` command
   (`apps/gridcraft-cli/src/check.rs`): how the cached values are compared.

## 2. What may be added

- **Only files generated here.** Every workbook is created from scratch by
  `tools/excel-oracles/generate.ps1` from a case in `tools/excel-oracles/cases.ps1`, with
  Microsoft Excel as the authoring tool. Never add a hand-made, downloaded, converted or found
  workbook, and never copy cells, formulas or values out of a real workbook. A failure found in a
  real file is rebuilt here as a synthetic shape (different numbers, no labels, no sheet names
  from the source).
- **No personal data** in contents, metadata, sheet names or file names. `verify-metadata.py`
  must pass on every file before a commit. It reads the local user name, computer name and
  profile from the environment; never write those (or anyone's name or e-mail) into a script.
- **Nothing copyrighted we don't own.** No Excel templates, themes beyond Excel's default, add-ins
  or sample files.

## 3. Adding or regenerating files, step by step

1. **Edit `cases.ps1`**: add a case whose path names the failure class (folder) and the shape and
   parameters (file name). Hyphens only, no underscores, in every file and folder name. Keep the
   file small: a few cells, one feature.
2. **Prerequisites:** Windows, Microsoft Excel installed, Python 3 for the metadata check
   (`GC_PYTHON=<path to python.exe>` if it isn't on PATH). Close nothing: the generator starts its
   own hidden Excel and only ever stops that one.
3. **Run** `powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-oracles/generate.ps1 '<regex>' -Force`
   for the cases you touched. Every line must read `ok` or `skip`, and the metadata line must say
   `0 problems`. Open a file or two in Excel and look.
4. **Check against GridCraft:** `gridcraft-cli check excel/ --json` (and the `.xls` files by
   name). New cases are expected to mismatch until GridCraft is fixed; that's the point. If one
   file crashes the engine (the `dependency-chain/` files do on 0.6.1), the whole run stops with
   no output, so check that file on its own.
5. **Always update `SHA256SUMS`** in the same commit:
   `find excel -type f \( -name '*.xlsx' -o -name '*.xls' \) | LC_ALL=C sort | xargs sha256sum --text > SHA256SUMS`
6. **One commit per batch**, naming what changed and the Excel version and build (the
   generator prints it on its first line).

## 4. Rules

- **Never rewrite published history:** no force-push, no amending pushed commits. Consumers pin
  commits, and every old pin must stay fetchable.
- **Don't regenerate unchanged cases.** Excel writes fresh timestamps and ids on every save, so a
  regenerated file always changes its hash. The generator skips existing files unless `-Force`.
- **Name time- and locale-dependent cases.** If Excel's answer depends on the current year, the
  locale or a system setting, put that in the file name (`-y2026`) or the README.
- **Keep the set small:** the whole of `excel/` should stay well under 10 MB.
- **One folder per source application,** with its generator under `tools/`.
