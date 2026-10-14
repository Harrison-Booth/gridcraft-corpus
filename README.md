# GridCraft Corpus

Real-file test oracles for [GridCraft](https://github.com/storytold/gridcraft): small workbooks
made in Microsoft Excel by a committed script. Each file carries Excel's own result next to
every formula (the cached `<v>` beside `<f>`), and GridCraft's `check` recalculates the file and
compares, cell by cell. Every file is generated from scratch; nothing comes from a real workbook.

**Contents:** [What this is](#what-this-is) · [Layout](#layout) · [Checking with GridCraft](#checking-with-gridcraft) ·
[Verifying](#verifying) · [Adding or regenerating files](#adding-or-regenerating-files) ·
[`excel/`](#excel-excel-oracle-workbooks) · [Metadata and privacy](#metadata-and-privacy) ·
[artcraft-fixtures](#relation-to-knuffixxartcraft-fixtures) · [Provenance](#provenance) · [License](#license)

## What this is

A workbook saved by Excel stores the value Excel computed for each formula. That cached value is
the oracle: GridCraft opens the file, recalculates every formula from scratch and reports each
cell whose result differs. The files are kept outside the GridCraft repository so its history
stays small, and a consumer should fetch them at a pinned commit and check them against
`SHA256SUMS`, so results can't change without a reviewed pin bump.

Every file covers one feature. Its path names the failure class and the file names the shape and
its parameters (`power-negative/caret-odd-roots.xlsx`), so a failing file says what broke.

## Layout

One top-level folder per authoring application:

| Folder | Source | Files |
|---|---|---:|
| [`excel/`](#excel-excel-oracle-workbooks) | Microsoft Excel for Microsoft 365 (16.0, build 20430, en-US), driven by [`tools/excel-oracles/`](tools/excel-oracles) | 47 |

Also in the repository:

- `tools/excel-oracles/`: the generator (`generate.ps1`), the case list (`cases.ps1`) and the
  metadata check (`verify-metadata.py`).
- `SHA256SUMS`: one `sha256  path` line for every corpus file.
- `LICENSE-MIT`, `LICENSE-APACHE`, `NOTICE`: the dual licence.

Third-party corpora are not copied here. This repository holds only files we generate.

## Checking with GridCraft

From a GridCraft checkout, build the CLI and point it at the folder:

```sh
cargo build --release -p gridcraft-cli
target/release/gridcraft-cli check ../gridcraft-corpus/excel --json
# check collects only .xlsx/.xlsm from a folder; name the .xls files explicitly:
target/release/gridcraft-cli check ../gridcraft-corpus/excel/legacy-xls/*.xls ../gridcraft-corpus/excel/precision-displayed/*.xls --json
```

`check` reports calc mismatches (expected = Excel's cached value) and a round-trip comparison
(GridCraft's saved copy against the source). Its default tolerance is 1e-9, relative above
magnitude 1 and absolute below, so a raw residue such as 1.2e-10 where Excel shows 0 passes on
its own and shows up only in the cells downstream (see `near-zero/`).

The upward chains in `dependency-chain/` overflow the stack of GridCraft 0.6.1 (main at
84acf62), which ends the whole `check` run with no output. Until that is fixed, check that folder
file by file, or leave it out.

## Verifying

```sh
sha256sum -c SHA256SUMS          # Git Bash or Linux; on macOS: shasum -a 256 -c SHA256SUMS
```

## Adding or regenerating files

1. **Only generator-made files.** Every file is written by `tools/excel-oracles/generate.ps1`
   from a case in `cases.ps1`. No hand-made, downloaded or found workbooks, and nothing from
   anyone's disk (see [`AGENTS.md`](AGENTS.md)).
2. **Add a case** to `cases.ps1`: a path under `excel/`, a list of cells (address, number, text
   or formula, optional number format) and the workbook settings (iteration, precision as
   displayed, `.xlsx` and/or `.xls`). A `build` step can add what cells can't (a PivotTable, a
   chart, a deleted sheet).
3. **Run the generator** on Windows with Excel installed:

   ```sh
   powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-oracles/generate.ps1               # every missing file
   powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-oracles/generate.ps1 '^date-text/' -Force   # a subset, regenerated
   ```

   The argument is a regular expression on the path (`near-zero/sub-reordered-parts-small.xlsx`).
   It prints one line per file: `ok` (generated), `skip` (already there; `-Force` regenerates)
   or `FAIL`. At the end it runs `verify-metadata.py` over `excel/`; any finding fails the run.
   Set `GC_PYTHON` if `python` isn't on PATH.
4. **Refresh the manifest** in the same commit:

   ```sh
   find excel -type f \( -name '*.xlsx' -o -name '*.xls' \) | LC_ALL=C sort | xargs sha256sum --text > SHA256SUMS
   ```

5. **One commit per batch**, naming what changed and the Excel version and build.

Excel never writes the same bytes twice (timestamps, document ids), which is why existing files
are skipped by default: regenerate a file only when its case changes. Never rewrite this
repository's history once it is published; consumers pin commits.

## `excel/`: Excel oracle workbooks

47 files (44 `.xlsx`, 3 `.xls`), 0.67 MB in total. The cases are synthetic shapes of failure
classes found by running GridCraft over real-world workbooks; only the shapes were kept.
Each file with a handful of formulas also has a text listing (`C1: =A1-B1`) two columns right of
its used range, so it explains itself when opened.

| Directory | Files | What it exercises | Excel's answer |
|---|---:|---|---|
| `near-zero/` | 7 | `=A1-B1` where A1 and B1 are the same three parts summed in a different order (`0.1+0.2+0.3`, `1000000.1+0.7+0.2`, and `300000000.1+0.1+0.1`, whose 5.96e-8 residue is above `check`'s tolerance); SUM variants; the same difference in parentheses, inside `IF(...>0)` and as a divisor; one copy with iteration on | The final `-` of a formula snaps to 0, so `=A1-B1` is 0 and `=1/C1` is `#DIV/0!`. Inside parentheses, a comparison or `SUM(A1-B1)` the residue (1.16e-10) survives, so `IF(A1-B1>0,…)` is TRUE |
| `power-negative/` | 4 | `^` and `POWER` with a negative base: odd roots (`(-8)^(1/3)`, `(-32)^(1/5)`, `(-8)^0.2`, negative exponents, cell operands), even roots and 14-digit exponents that stay `#NUM!`, and the exponent limit (1/2147483645 works, 1/2147483647 and 1/1E15 don't) | `(-8)^(1/3)` = -1.9999999999999998 |
| `criteria-ref-error/` | 10 | `#REF!` in a range argument of SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS; a `#REF!` criterion as a control; the realistic origin, a deleted sheet | `#REF!` for every range position; a `#REF!` criterion gives 0 |
| `date-text/` | 5 | Date text with month names and without separators: `"1Jan"`, `"Jan1"`, `"1 Jan"`, `"Jan 2021"`, `"1Jan21"`, `"15Mar2024"`, `MONTH(1&"Feb")`, `DAY("15Mar")`; numbers read as two-digit years (`"Jan 32"`, `"Feb 29"`); inputs that stay `#VALUE!` | See the year note below |
| `precision-displayed/` | 3 | Precision as displayed on: `=-130+657.91` formatted `0`, `=A1*10`, `=A1/7` formatted `0.0000`, `=A3*7`, `=-1+2.91`, in `.xlsx` and `.xls`; the same cells with the option off as a control | 528, 5280, 75.4286, 528.0002 |
| `legacy-xls/` | 4 | Legacy (pre-dynamic-array) formulas with a range in one cell, `=Sheet1!A1:G1` and `=Sheet1!A1:A7`, saved as Excel 97-2003 `.xls` and, as a twin, `.xlsx` | Implicit intersection: the cell in the formula's column (row); `#VALUE!` outside the range |
| `circular/` | 7 | `A1 =SUMIFS(B1:C1,B2:C2,"x")` with `C1 =A1*2`: C1 not selected by the criteria (not circular), selected (truly circular, iteration off and on), in a later criteria range, whole columns; the SUMIF form; INDEX as a control | Not circular unless the criteria select C1; the truly circular file caches A1 = 5, C1 = 0 |
| `dependency-chain/` | 3 | `=A{r+1}+1` down 5,000 rows (each cell reads the row below), a 4,000-link chain hopping between two sheets, and the downward chain as a control. Excel stores each column as a shared formula | 5000 and 4000 at the top |
| `rate/` | 2 | `RATE(n,0,pv,0)` and `RATE(n,0,pv,tiny fv)` | `RATE(5,0,-100,0)` = -0.9999994091451825; `RATE(1,0,-100,0)` and `RATE(5,0,0,0)` are `#NUM!` |
| `round-trip/` | 2 | A PivotTable from an in-sheet range (row field, sum of values) with `GETPIVOTDATA`, and a clustered column chart over a range | For round-trip loss; their formulas are a light calc check |

**Locale.** Excel ran as en-US (country code 1, decimal `.`, month-day-year). Date-text parsing
depends on it, and on Windows' two-digit-year window (1930-2029): `"Feb 29"` is February 2029 and
`"Feb 30"` February 1930, not a day in the current year.

**Year.** Date text without a year takes the year the calculation runs in. Only
`date-text/month-name-no-year-y2026.xlsx` depends on it: its cached serials are right during
2026. From 2027 on, a correct engine will mismatch every cell in that file; regenerate it with
the new year in its name. `MONTH`/`DAY` extractions don't depend on the year.

**RATE.** Excel's RATE is an iterative solver, so its answer for a degenerate case (pmt 0, fv 0)
is wherever its iteration stopped, about -0.9999994. A correct-in-spirit implementation may not
match to 1e-9.

**Iteration.** `circular/sumifs-selected-cell-iterate-on.xlsx` and
`near-zero/downstream-if-positive-iterate-on.xlsx` are saved with iterative calculation on (100
iterations, max change 0.001). The circular one converges towards 10 and stops where Excel's
iteration stopped, so a near match is the expected result there.

## Metadata and privacy

- The generator saves every workbook with Excel's own **remove personal information on save**
  (`Workbook.RemovePersonalInformation`). Excel for Microsoft 365 otherwise writes the signed-in
  Office identity into `docProps/core.xml` (`lastModifiedBy`) and the `.xls` WRITEACCESS record,
  whatever `Application.UserName` says, and the save folder into `xl/workbook.xml`
  (`x15ac:absPath`). With the option on, Author, Last Author and Company are saved empty,
  `absPath` is left out and `workbookPr` carries `filterPrivacy="1"`.
- Files are saved through a neutral staging folder (`C:\gc-corpus-stage`) and then moved.
- `verify-metadata.py` checks every file: the author fields are empty (or "GridCraft
  contributors"), there is no `docProps/custom.xml`, and no part (or, for `.xls`, no byte run in
  Latin-1 or UTF-16LE) contains a user folder or the generating machine's user name, display
  name, computer name or domain. It reads those from the environment and from Excel, so they are
  never written into this repository.
- No Excel setting is changed for good. Calculation settings (iteration, precision as
  displayed) are set for every case; `SheetsInNewWorkbook` is put back at the end.

## Relation to KNUFFIXX/artcraft-fixtures

[KNUFFIXX/artcraft-fixtures](https://github.com/KNUFFIXX/artcraft-fixtures) holds small
synthetic files behind bug reports to the ArtCraft projects, with a `gridcraft/` folder of
openpyxl-written workbooks whose expected values (measured in a de-DE Excel through COM) are
written into cells next to the test cells. Where it is cheap we follow its conventions, and the
two could be merged later:

- **Same shape:** one folder per target, every file written by a committed script in the
  repository, no real-world content, MIT OR Apache-2.0.
- **Self-describing files:** like their expected-value cells, our files carry a text listing of
  their formulas next to the used range. Our expected values are Excel's own cached results,
  which `gridcraft-cli check` reads directly, so no expected-value cells are needed.
- **README table:** the `excel/` table above has their File / Reproduces columns; a `Report`
  column can be added as issues are filed.
- **Differences:** hyphens instead of `gc_` underscores in names, a subfolder per failure class,
  the authoring application (Excel itself) instead of openpyxl, en-US instead of de-DE, and a
  `SHA256SUMS` manifest for pinning. To contribute a file there, copy it under
  `gridcraft/gc_<name>.xlsx` and add the case's formula to their table.

## Provenance

Everything here is original work by the GridCraft corpus contributors. The files are generated
from scratch by the scripts in `tools/`, with Microsoft Excel used only as the authoring and
calculating tool. No Excel templates, add-ins or sample files are included. Every number, text
and formula is synthetic.

## License

Dual-licensed under [MIT](LICENSE-MIT) or [Apache-2.0](LICENSE-APACHE), at your option.
Copyright (c) 2026 the GridCraft corpus contributors. See [NOTICE](NOTICE).

<sub>Microsoft and Excel are trademarks of the Microsoft group of companies. This project is not
affiliated with, sponsored by or endorsed by Microsoft; the names are used only to describe the
files' origin.</sub>
