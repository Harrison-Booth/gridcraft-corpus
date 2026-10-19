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
| [`excel/`](#excel-excel-oracle-workbooks) | Microsoft Excel for Microsoft 365 (16.0, build 20430, en-US), driven by [`tools/excel-oracles/`](tools/excel-oracles) | 107 |

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

107 files (101 `.xlsx`, 6 `.xls`), 1.53 MB in total. The cases are synthetic shapes of failure
classes found by running GridCraft over real-world workbooks; only the shapes were kept.
Each file with a handful of formulas also has a text listing (`C1: =A1-B1`) two columns right of
its used range, so it explains itself when opened.

| Directory | Files | What it exercises | Excel's answer |
|---|---:|---|---|
| `near-zero/` | 7 | `=A1-B1` where A1 and B1 are the same three parts summed in a different order (`0.1+0.2+0.3`, `1000000.1+0.7+0.2`, and `300000000.1+0.1+0.1`, whose 5.96e-8 residue is above `check`'s tolerance); SUM variants; the same difference in parentheses, inside `IF(...>0)` and as a divisor; one copy with iteration on | The final `-` of a formula snaps to 0, so `=A1-B1` is 0 and `=1/C1` is `#DIV/0!`. Inside parentheses, a comparison or `SUM(A1-B1)` the residue (1.16e-10) survives, so `IF(A1-B1>0,…)` is TRUE |
| `power-negative/` | 4 | `^` and `POWER` with a negative base: odd roots (`(-8)^(1/3)`, `(-32)^(1/5)`, `(-8)^0.2`, negative exponents, cell operands), even roots and 14-digit exponents that stay `#NUM!`, and the exponent limit (1/2147483645 works, 1/2147483647 and 1/1E15 don't) | `(-8)^(1/3)` = -1.9999999999999998 |
| `criteria-ref-error/` | 10 | `#REF!` in a range argument of SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS; a `#REF!` criterion as a control; the realistic origin, a deleted sheet | `#REF!` for every range position; a `#REF!` criterion gives 0 |
| `date-text/` | 11 | Date text with month names and without separators: `"1Jan"`, `"Jan1"`, `"1 Jan"`, `"Jan 2021"`, `"1Jan21"`, `"15Mar2024"`, `MONTH(1&"Feb")`, `DAY("15Mar")`; numbers read as two-digit years (`"Jan 32"`, `"Feb 29"`), also with `-` and `/`; `-`/`/` day forms (`"Jan-21"`, `"Jan/1"`); the two-digit-year window (`"1Jan29"`, `"1Jan30"`); month prefixes longer than three letters (`"Sept"`, `"Janu"`, `"Decemb"`); punctuated forms that stay `#VALUE!` (`"4.Mar.2026"`, `"1Jan, 2021"`, `"Jan 15 2021"`, `"001Jan"`, `"Ma"`) and those that parse (`"Mar 4, 2026"`) | `"Jan-21"` is 21 January of the current year, not 2021; `"1Jan29"` is 2029 and `"1Jan30"` 1930. See the year note below |
| `precision-displayed/` | 16 | Precision as displayed on: `=-130+657.91` formatted `0`, `=A1*10`, `=A1/7` formatted `0.0000`, `=A3*7`, `=-1+2.91`, in `.xlsx` and `.xls`; the same cells with the option off as a control. One file per format family (percent, currency, `0.0,,` scaling, fractions, scientific, `0.00;(0.0)` sections, literal text, `[<1]0.00;0`), with General, date/time and `@` as controls; a cross-cell read; a spill landing in cells formatted `0`, `0.0`, `0.00` | 528, 5280, 75.4286, 528.0002. Every number format rounds to what it shows (`0.0,,` stores 1234567.891 as 1200000, `?/8` stores 0.3 as 0.25); dates, times, General and `@` keep the full value. Each spilled element is rounded by its own cell's format |
| `legacy-xls/` | 10 | Legacy (pre-dynamic-array) formulas with a range in one cell, `=Sheet1!A1:G1` and `=Sheet1!A1:A7`; a range inside an operator inside SUM, `=MAX(SUM(A1:A3/12),0)`; a one-cell Ctrl+Shift+Enter array; a formula filled down (a shared formula). Each saved as Excel 97-2003 `.xls` and, as a twin, `.xlsx` | Implicit intersection: the cell in the formula's column (row); `#VALUE!` outside the range. `=MAX(SUM(A1:A3/12),0)` is 2 in row 2 and 6 as an array formula |
| `circular/` | 16 | `A1 =SUMIFS(B1:C1,B2:C2,"x")` with `C1 =A1*2`: C1 not selected by the criteria (not circular), selected (truly circular, iteration off and on), in the first or a later criteria range, whole columns, size-mismatched ranges, an array criterion that does or doesn't select C1; the SUMIF, AVERAGEIFS, MAXIFS, MINIFS and COUNTIFS forms; INDEX as a control | Not circular unless the criteria select C1, or C1 is in the first criteria range (SUMIFS and COUNTIFS); size-mismatched ranges are `#VALUE!`, not circular. The truly circular files cache C1 = 0 |
| `dependency-chain/` | 3 | `=A{r+1}+1` down 5,000 rows (each cell reads the row below), a 4,000-link chain hopping between two sheets, and the downward chain as a control. Excel stores each column as a shared formula | 5000 and 4000 at the top |
| `rate/` | 2 | `RATE(n,0,pv,0)` and `RATE(n,0,pv,tiny fv)` | `RATE(5,0,-100,0)` = -0.9999994091451825; `RATE(1,0,-100,0)` and `RATE(5,0,0,0)` are `#NUM!` |
| `round-trip/` | 3 | A PivotTable from an in-sheet range (row field, sum of values) with `GETPIVOTDATA`, a clustered column chart over a range, and a two-cell workbook whose author fields are blank but whose `docProps/core.xml` carries the created and modified dates | For round-trip loss; their formulas are a light calc check |
| `pivot-external/` | 4 | PivotTables whose source is in a second workbook, `source-data.xlsx` (committed next to them): a range, a defined name in that workbook, and two PivotTables sharing one cache. The link is saved relative (`Target="source-data.xlsx"`), so it resolves inside the repository | `GETPIVOTDATA` over them gives 58, 392, 156 (and 107 for the second PivotTable) |
| `getpivotdata/` | 12 | `GETPIVOTDATA` against an in-sheet PivotTable with two row fields: grand total, one and two field/item pairs, the data field by caption, mixed case, a cell inside the PivotTable as the reference, an item from a cell; a missing item, an inner field alone and an unknown data field; SUMIFS over the same rows as a control; ten more shapes: row and column fields, report filters, totals off, dates and numbers as items, renamed captions, hidden items, several data fields, a stale report (source edited, PivotTable not refreshed) and a tabular PivotTable on another sheet | 392, 107, 72; `#REF!` for the missing item, the inner field alone and the unknown field |
| `pivot-data-model/` | 3 | PivotTables whose cache is a workbook connection or the Data Model, not a worksheet range: a worksheet table (`Sales`) loaded into the model; an OLE DB connection to a workbook the generator writes into its staging folder (not in the model); the same OLE DB connection loaded into the model. Each has `GETPIVOTDATA` formulas next to it | `cacheSource type="external"` with a `connectionId` and `xl/connections.xml`; the two model files also carry `xl/model/item.data`. `GETPIVOTDATA` gives 392, 107, 166 (model forms use `[Measures].[Sum of Amount]` and `[Sales].[Region].&[north]`; the plain `"Sum of Amount"` and a bare item are `#REF!` on the model) |

**Locale.** Excel ran as en-US (country code 1, decimal `.`, month-day-year). Date-text parsing
depends on it, and on Windows' two-digit-year window (1930-2029): `"Feb 29"` is February 2029 and
`"Feb 30"` February 1930, not a day in the current year.

**Year.** Date text without a year takes the year the calculation runs in. Only
`date-text/month-name-no-year-y2026.xlsx` and `date-text/month-name-day-this-year-y2026.xlsx`
depend on it: their cached serials are right during 2026. From 2027 on, a correct engine will
mismatch every cell in those files; regenerate them with the new year in their names. `MONTH`/`DAY` extractions don't depend on the year.

**RATE.** Excel's RATE is an iterative solver, so its answer for a degenerate case (pmt 0, fv 0)
is wherever its iteration stopped, about -0.9999994. A correct-in-spirit implementation may not
match to 1e-9.

**Pivot sources.** The `pivot-external/` files read their source from `source-data.xlsx` in the
same folder. Keep the four together; regenerating a pivot file needs `source-data.xlsx` to exist
first (the generator copies it into its staging folder and opens it from there, which is what
makes Excel write the link relative).

**Iteration.** `circular/sumifs-selected-cell-iterate-on.xlsx` and
`near-zero/downstream-if-positive-iterate-on.xlsx` are saved with iterative calculation on (100
iterations, max change 0.001). The circular one converges towards 10 and stops where Excel's
iteration stopped, so a near match is the expected result there.

### Files added in the second batch

| File | Reproduces |
|---|---|
| `pivot-external/source-data.xlsx` | The source workbook: 8 rows of Region, Category, Amount and a workbook-level name `SourceRows` over them. No formulas |
| `pivot-external/range-source.xlsx` | PivotTable whose cache reads `[source-data.xlsx]Sheet1!A1:C9`; `GETPIVOTDATA` against it |
| `pivot-external/defined-name-source.xlsx` | The same PivotTable with the cache reading the defined name `source-data.xlsx!SourceRows` |
| `pivot-external/shared-cache.xlsx` | Two PivotTables (by Category, by Region) on one external cache |
| `getpivotdata/field-item-forms.xlsx` | Ten `GETPIVOTDATA` forms, six with values and four that are `#REF!` |
| `getpivotdata/sumifs-control.xlsx` | Control: the same PivotTable, totals from `SUM`/`SUMIFS` only |
| `getpivotdata/row-col-fields.xlsx` | Region by Category cross-tab: grand total, one and two items in either order, an empty intersection (0), a cell inside the data area as the reference; a reference off the PivotTable and an unknown item are `#REF!` |
| `getpivotdata/report-filter.xlsx` | A report filter showing one item, one showing (All), and one with a single item hidden: the totals follow what is shown, and items filtered out are `#REF!` |
| `getpivotdata/totals-off.xlsx` | Subtotals and grand totals turned off: `#REF!` for the grand total and for Region alone, values for Region with Category |
| `getpivotdata/date-number-items.xlsx` | Dates and numbers as items: `DATE()`, a serial and the text `3/10/2026` give values, `2026-01-15` and a date with no data are `#REF!`; the numbers `10` and `"20"` give values, `10.5` is `#REF!` |
| `getpivotdata/captions.xlsx` | A renamed data field, field and item: the caption (`Area`, `North Zone`) is what Excel accepts for the field and the item; the data field answers to its caption and to its source name, and the source field and item names are `#REF!` |
| `getpivotdata/hidden-items.xlsx` | Items hidden by a filter (a Region and a Category): a hidden item is `#REF!` and the grand total covers what is shown |
| `getpivotdata/multi-data-fields.xlsx` | Four data fields (sum, count, average, % of total), each asked for by its caption; the plain source name answers with the first |
| `getpivotdata/stale-report.xlsx` | Source edited and the PivotTable not refreshed: `GETPIVOTDATA` returns the displayed (old) value, `SUMIFS` the new one |
| `getpivotdata/stale-report-sorted.xlsx` | The same with the source sorted by Region, which separates a stale report from an item-order difference |
| `getpivotdata/tabular-other-sheet.xlsx` | A tabular PivotTable on another sheet: a field/item pair, a missing field and a reference given as a wide range |
| `date-text/month-name-day-this-year-y2026.xlsx` | `"Jan-21"`, `"Jan/21"`, `"Jan-1"`, `"Jan/1"`, `"Jan-30"`, `"1Janu"`, `"Janu 1"`, `"1Febr"`, `"01Jan"`, `"Sep 1"`: days in the current year |
| `date-text/month-name-number-as-year-separators.xlsx` | `"Jan-32"`, `"Jan/32"`, `"Jan-0"`, `"Feb-30"` and `"Jan 32"+0`-style coercion: the number is a year |
| `date-text/two-digit-year-window.xlsx` | `"1Jan29"` is 2029, `"1Jan30"` 1930, with and without separators |
| `date-text/month-name-long-prefix.xlsx` | `"Sept"`, `"Janu"`, `"Septe"`, `"Septem"`, `"Decemb"` with a year parse; `"Janux 2021"` and `"Sept"` alone are `#VALUE!` |
| `date-text/month-name-punctuation-value-errors.xlsx` | `"4.Mar.2026"`, `"Jan.21"`, `"1Jan, 2021"`, `"4 Mar, 2026"`, `"2021Jan"`, `"Jan 15 2021"`, `"001Jan"`, `"Ma"`, `"Sept 1 2021"`, `"Mar.4"`: all `#VALUE!` |
| `date-text/month-name-punctuation-control.xlsx` | Control: `"Mar 4, 2026"`, `"Jan 15, 2021"`, `"4-Mar-2026"`, `"4/Mar/2026"`, `"Jan-2021"`, `"Jan/2021"` parse |
| `precision-displayed/pad-on-percent.xlsx` | `0%`, `0.0%`, `0.00%`: `=1/3` is stored as 0.33 |
| `precision-displayed/pad-on-currency.xlsx` | `$#,##0.00`, a negative section in parentheses, `"$"#,##0.0` |
| `precision-displayed/pad-on-thousands-scaling.xlsx` | `0.0,,`, `#,##0,`, `0.00,`: rounded at the scaled digit (1200000, 12000, 9880) |
| `precision-displayed/pad-on-fractions.xlsx` | `# ?/?` and `?/8`: stored as the fraction shown (0.3 is 2/7 and 2/8) |
| `precision-displayed/pad-on-scientific.xlsx` | `##0.0E+0`, `0.00E+00`, `0.0E+0` |
| `precision-displayed/pad-on-multi-section.xlsx` | `0.00;(0.0)` and `0;-0.000`: the negative section's decimals apply to negatives |
| `precision-displayed/pad-on-literal-text.xlsx` | `0.0" u"`, `"x"0.00`, and `0.0"%"` (a quoted `%` doesn't scale) |
| `precision-displayed/pad-on-conditional-sections.xlsx` | `[<1]0.00;0`: rounded to two decimals whatever the section shows (`=2+1/3` shows 2, stores 2.33; `=-10/3` stores -3) |
| `precision-displayed/pad-on-general-control.xlsx` | Control: General keeps the full value |
| `precision-displayed/pad-on-date-time-control.xlsx` | Control: `m/d/yyyy`, `h:mm`, `m/d/yyyy h:mm`, `[h]:mm` keep the full value |
| `precision-displayed/pad-on-text-format-control.xlsx` | Control: a formula in a cell formatted `@` (set after entry) keeps the full value |
| `precision-displayed/pad-on-cross-cell-read.xlsx` | `=A1` formatted `0.0` reading a General `=10/3` stores 3.3; a General cell reading a formatted one reads the rounded value |
| `precision-displayed/pad-on-spill-mixed-formats.xlsx` | `=SEQUENCE(3,1,1/3,1/3)` spilling into cells formatted `0`, `0.0`, `0.00`: stored as 0, 0.7, 1; `SUM` over them is 1.7 |
| `circular/sumifs-first-criteria-range.xlsx` | C1 in the first criteria range: truly circular (C1 cached 0); the twin of `sumifs-later-criteria-range.xlsx` |
| `circular/sumifs-size-mismatch-value-error.xlsx` | Criteria range one cell wider than the sum range: `#VALUE!`, not circular |
| `circular/sumifs-array-criterion-unselected.xlsx` | `=SUM(SUMIFS(B1:C1,B2:C2,{"x","z"}))`: C1 not selected, not circular |
| `circular/sumifs-array-criterion-selected.xlsx` | `{"x","y"}` selects C1: truly circular |
| `circular/averageifs-unselected-cell.xlsx` | AVERAGEIFS, C1 not selected: not circular |
| `circular/maxifs-unselected-cell.xlsx` | MAXIFS, C1 not selected: not circular |
| `circular/minifs-unselected-cell.xlsx` | MINIFS, C1 not selected: not circular |
| `circular/countifs-later-criteria-range.xlsx` | COUNTIFS with C1 in the second criteria range: not circular |
| `circular/countifs-first-criteria-range.xlsx` | COUNTIFS with C1 in the first criteria range: truly circular |
| `legacy-xls/cse-single-cell-array.xls`, `.xlsx` | One-cell Ctrl+Shift+Enter arrays `{=SUM(A1:A3*B1:B3)}`, `{=MAX(A1:A3-B1:B3)}` stay arrays (32, -3); SUMPRODUCT as a control |
| `legacy-xls/shared-formula-fill-down.xls`, `.xlsx` | Two formulas filled down six rows, stored as shared formulas (SHRFMLA in `.xls`) |
| `legacy-xls/implicit-intersection-inside-sum.xls`, `.xlsx` | `=MAX(SUM(A1:A3/12),0)` as a legacy formula: 2 in row 2, `#VALUE!` in row 5; 6 as an array formula |
| `round-trip/core-properties-blank-author.xlsx` | Author and Last Author blank, created and modified dates present: a save must keep `docProps/core.xml` |

### Files added in the third batch

| File | Reproduces |
|---|---|
| `pivot-data-model/model-from-table.xlsx` | Worksheet table `Sales` (Region, Category, Amount; 8 rows) added to the Data Model; a PivotTable on the model (Region rows, sum of Amount) on a second sheet. Five `GETPIVOTDATA` formulas: grand total 392, `north` 107, `south` 166, and two `#REF!` forms |
| `pivot-data-model/oledb-connection.xlsx` | OLE DB connection (ACE provider, type 5) to a generated workbook, not in the model; a PivotTable whose cache reads the connection, with records saved. `GETPIVOTDATA` gives 392, 107, 166. The source path is `C:\gc-corpus-stage\source-data.xlsx`, which does not exist after the run |
| `pivot-data-model/oledb-connection-in-model.xlsx` | The same OLE DB connection loaded into the Data Model (type 100 connection plus `ThisWorkbookDataModel`); a PivotTable on the model. `GETPIVOTDATA` gives 392 and 107 |

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
