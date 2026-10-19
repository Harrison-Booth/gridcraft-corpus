# The case list for generate.ps1. Dot-sourced; returns an array of case hashtables.
#
# A case:
#   name      path under excel/ without extension, e.g. 'near-zero/sub-reordered-parts-small'
#   cells     list of cell specs (helpers below):
#               F 'A1' '=A2+1' ['0.00']   formula (Range.Formula: legacy, implicit intersection)
#               V 'A1' 1.5 ['0.00']       number constant
#               T 'A1' 'x'                text constant (cell formatted as text first)
#               @{ a = 'A1'; f2 = '=SEQUENCE(3)' }        dynamic-array formula (Range.Formula2)
#               @{ a = 'A1'; fa = '=SUM(A1:A3*B1:B3)' }   Ctrl+Shift+Enter array (Range.FormulaArray)
#               @{ a = 'A1'; n = '0.0' }                 number format only
#             An address may carry a sheet ('Sheet2!A1') and may be a range
#             ('A1:A4999'): a formula with relative references then fills the range.
#   sheets    sheet names, default @('Sheet1')
#   formats   @('xlsx'), @('xls') or both; default @('xlsx')
#   iterate   iterative calculation on (default $false); maxIter / maxChange go with it
#   pad       precision as displayed (default $false)
#   build     optional scriptblock { param($wb, $xl, $ctx) ... } run after the cells, before
#             saving; $ctx.stage is the staging folder and $ctx.outRoot the excel/ folder
#   list      write a formula listing ("C1: =A1-B1" text cells) to the right of the used
#             range on each sheet, so a file explains itself when opened (default $true)
#
# Every number, text and formula here is synthetic. Nothing comes from a real workbook.

function F([string]$a, [string]$f, [string]$n = $null) { @{ a = $a; f = $f; n = $n } }
function V([string]$a, [double]$v, [string]$n = $null) { @{ a = $a; v = $v; n = $n } }
function T([string]$a, [string]$t) { @{ a = $a; t = $t } }

# Formula cells in A1, A2, ... for a list of formulas.
function Column([string[]]$formulas, [string]$col = 'A') {
    $i = 0
    foreach ($f in $formulas) { $i++; F "$col$i" $f }
}

$cases = @()

# ---------------------------------------------------------------- 1. near-zero snapping
# The same parts summed in two orders differ in the last bit. Excel snaps a tiny result of
# the formula's final + or - to 0, but not inside parentheses or a comparison.
$small = @((V 'D1' 0.1), (V 'D2' 0.2), (V 'D3' 0.3))
$large = @((V 'D1' 1000000.1), (V 'D2' 0.7), (V 'D3' 0.2))
$subShape = @(
    (F 'A1' '=D1+D2+D3'), (F 'B1' '=D3+D2+D1'),
    (F 'C1' '=A1-B1'), (F 'C2' '=(A1-B1)'), (F 'C3' '=(A1-B1)*1'), (F 'C4' '=B1-A1')
)
$cases += @{ name = 'near-zero/sub-reordered-parts-small'; cells = $small + $subShape }
$cases += @{ name = 'near-zero/sub-reordered-parts-large'; cells = $large + $subShape }
# Residue 5.96e-8: above gridcraft-cli check's default 1e-9 tolerance, so the root snap itself
# shows as a mismatch (the residues above pass the tolerance and show only downstream).
$cases += @{ name = 'near-zero/sub-reordered-parts-above-tolerance'; cells = @((V 'D1' 300000000.1), (V 'D2' 0.1), (V 'D3' 0.1)) + $subShape }
$cases += @{ name = 'near-zero/sum-reordered-parts-large'; cells = $large + @(
        (F 'A1' '=SUM(D1:D3)'), (F 'B1' '=SUM(D3,D2,D1)'),
        (F 'C1' '=SUM(D1:D3)-SUM(D3,D2,D1)'), (F 'C2' '=SUM(A1,-B1)'), (F 'C3' '=SUM(A1-B1)'),
        (F 'C4' '=A1-B1'), (F 'C5' '=SUM(C4)')) }
# A1 (reverse order) is the larger one here, so the raw residue is positive.
$downstream = @(
    (F 'A1' '=D3+D2+D1'), (F 'B1' '=D1+D2+D3'), (F 'C1' '=A1-B1'),
    (F 'E1' '=IF(C1>0,1,0)'), (F 'E2' '=IF(A1-B1>0,1,0)'), (F 'E3' '=IF(C1=0,1,0)'),
    (F 'E4' '=IF((A1-B1)>0,1,0)')
)
$cases += @{ name = 'near-zero/downstream-if-positive'; cells = $large + $downstream }
$cases += @{ name = 'near-zero/downstream-if-positive-iterate-on'; cells = $large + $downstream; iterate = $true; maxIter = 100; maxChange = 0.001 }
$cases += @{ name = 'near-zero/downstream-division'; cells = $large + @(
        (F 'A1' '=D3+D2+D1'), (F 'B1' '=D1+D2+D3'), (F 'C1' '=A1-B1'),
        (F 'E1' '=1/C1'), (F 'E2' '=1/(A1-B1)'), (F 'E3' '=IFERROR(1/C1,-1)'), (F 'E4' '=ISERROR(1/C1)')) }

# ---------------------------------------------------------------- 2. odd roots of negatives
$cases += @{ name = 'power-negative/caret-odd-roots'; cells = @(Column @(
            '=(-8)^(1/3)', '=(-32)^(1/5)', '=(-27)^(1/3)', '=(-1)^(1/3)', '=(-8)^(1/9)', '=(-8)^0.2',
            '=(-8)^(-1/3)', '=(-8)^0.333333333333333', '=-8^(1/3)')) + @(
        (V 'C1' -8), (F 'C2' '=1/3'), (F 'C3' '=C1^C2')) }
$cases += @{ name = 'power-negative/power-function-odd-roots'; cells = @(Column @(
            '=POWER(-8,1/3)', '=POWER(-32,1/5)', '=POWER(-8,-1/3)', '=POWER(-8,1/9)', '=POWER(-8,0.2)')) + @(
        (V 'C1' -8), (F 'C2' '=1/3'), (F 'C3' '=POWER(C1,C2)')) }
$cases += @{ name = 'power-negative/even-roots-num-error'; cells = @(Column @(
            '=(-8)^(2/3)', '=(-8)^(1/4)', '=(-16)^(1/2)', '=POWER(-27,2/3)', '=POWER(-16,0.5)',
            '=(-8)^0.33333333333333', '=SQRT(-4)')) }
$cases += @{ name = 'power-negative/odd-root-exponent-limits'; cells = @(Column @(
            '=(-8)^(1/2147483645)', '=(-8)^(1/2147483647)', '=(-8)^(1/1E15)', '=POWER(-8,1/2147483645)')) }

# ---------------------------------------------------------------- 3. #REF! in criteria ranges
$crit = @((V 'A1' 1), (V 'A2' 2), (V 'A3' 1), (V 'B1' 10), (V 'B2' 20), (V 'B3' 30))
$refCases = [ordered]@{
    'sumif'      = @('=SUMIF(#REF!,1,B1:B3)', '=SUMIF(A1:A3,1,#REF!)', '=SUMIF(#REF!,1)')
    'sumifs'     = @('=SUMIFS(#REF!,A1:A3,1)', '=SUMIFS(B1:B3,#REF!,1)', '=SUMIFS(B1:B3,A1:A3,1,#REF!,1)')
    'countif'    = @('=COUNTIF(#REF!,1)', '=COUNTIF(#REF!,">0")')
    'countifs'   = @('=COUNTIFS(#REF!,1)', '=COUNTIFS(A1:A3,1,#REF!,1)')
    'averageif'  = @('=AVERAGEIF(#REF!,1,B1:B3)', '=AVERAGEIF(A1:A3,1,#REF!)', '=AVERAGEIF(#REF!,1)')
    'averageifs' = @('=AVERAGEIFS(#REF!,A1:A3,1)', '=AVERAGEIFS(B1:B3,#REF!,1)')
    'maxifs'     = @('=MAXIFS(#REF!,A1:A3,1)', '=MAXIFS(B1:B3,#REF!,1)')
    'minifs'     = @('=MINIFS(#REF!,A1:A3,1)', '=MINIFS(B1:B3,#REF!,1)')
    # Control: a #REF! criterion (not range) gives 0 in Excel.
    'criterion-control' = @('=SUMIF(A1:A3,#REF!,B1:B3)', '=COUNTIF(A1:A3,#REF!)', '=SUMIFS(B1:B3,A1:A3,#REF!)')
}
foreach ($k in $refCases.Keys) {
    $cases += @{ name = "criteria-ref-error/$k"; cells = $crit + @(Column $refCases[$k] 'D') }
}
# The realistic origin: the referenced sheet is deleted, and Excel rewrites the ranges.
$cases += @{ name = 'criteria-ref-error/deleted-sheet'; sheets = @('Sheet1', 'Sheet2')
    cells = @((V 'Sheet2!A1' 1), (V 'Sheet2!A2' 2), (V 'Sheet2!A3' 1), (V 'Sheet2!B1' 10), (V 'Sheet2!B2' 20), (V 'Sheet2!B3' 30),
        (F 'Sheet1!A1' '=SUMIFS(Sheet2!B1:B3,Sheet2!A1:A3,1)'), (F 'Sheet1!A2' '=COUNTIF(Sheet2!A1:A3,1)'),
        (F 'Sheet1!A3' '=SUMIF(Sheet2!A1:A3,1,Sheet2!B1:B3)'), (F 'Sheet1!A4' '=AVERAGEIFS(Sheet2!B1:B3,Sheet2!A1:A3,1)'))
    build = { param($wb, $xl) $wb.Worksheets.Item('Sheet2').Delete() } }

# ---------------------------------------------------------------- 4. date text with month names
# Locale: Excel en-US (country 1, M/D/Y). Text without a year takes the CURRENT year, so the
# cached serials in the -y2026 file are right only during 2026.
$cases += @{ name = 'date-text/month-name-no-year-y2026'; cells = @(Column @(
            '=DATEVALUE("1Jan")', '=DATEVALUE("15Mar")', '=DATEVALUE("Jan1")', '=DATEVALUE("Mar15")',
            '=DATEVALUE("1January")', '=DATEVALUE("1 Jan")', '=DATEVALUE("Sept 1")', '="1Jan"+0', '=(1&"Feb")+0')) }
$cases += @{ name = 'date-text/month-name-explicit-year'; cells = @(Column @(
            '=DATEVALUE("Jan 2021")', '=DATEVALUE("Jan2021")', '=DATEVALUE("1Jan21")', '=DATEVALUE("1Jan2021")',
            '=DATEVALUE("15Mar2024")', '="15Mar2024"+0', '=DATEVALUE("15 March 2024")', '=DATEVALUE("1 Jan 21")')) }
$cases += @{ name = 'date-text/month-name-part-extract'; cells = @(Column @(
            '=MONTH(1&"Feb")', '=MONTH(1&"Jan")', '=DAY("15Mar")', '=MONTH("Jan 2021")', '=YEAR("1Jan2021")',
            '=MONTH("15Mar2024")')) }
# A number after a month name that can't be a day is read as a two-digit year (1930-2029 window).
$cases += @{ name = 'date-text/month-name-number-as-year'; cells = @(Column @(
            '=DATEVALUE("Jan 32")', '=DATEVALUE("Jan 0")', '=DATEVALUE("Feb 29")', '=DATEVALUE("Feb 30")')) }
$cases += @{ name = 'date-text/month-name-value-errors'; cells = @(Column @(
            '=DATEVALUE("Jan")', '=DATEVALUE("March")', '=DATEVALUE("32Jan")', '=DATEVALUE("0Jan")',
            '=DATEVALUE("Foo 2021")', '=DATEVALUE("1Jxn")', '=DATEVALUE("2021 Jan")', '=DATEVALUE("1Jan2021x")',
            '=DATEVALUE("29Feb2023")', '=MONTH("Jan")')) }

# ---------------------------------------------------------------- 5. precision as displayed
# The sources are formulas: with the option on, constants are rounded once on entry, so a
# constant carries no signal. Formats are set before the formulas.
$pad = @(
    (F 'A1' '=-130+657.91' '0'), (F 'A2' '=A1*10'), (F 'A3' '=A1/7' '0.0000'), (F 'A4' '=A3*7'),
    (F 'A5' '=-1+2.91' '0'), (F 'A6' '=A5*10'), (F 'A7' '=SUM(A1,A5)' '0.0')
)
$cases += @{ name = 'precision-displayed/pad-on'; cells = $pad; pad = $true; formats = @('xlsx', 'xls') }
$cases += @{ name = 'precision-displayed/pad-off-control'; cells = $pad; pad = $false }

# ---------------------------------------------------------------- 6. legacy formulas in .xls
# Range.Formula enters a legacy formula, so a range in one cell is implicitly intersected
# (Formula2 reads back as =@Sheet1!A1:G1). The .xlsx twin is the same formula in the newer file.
$row = @(); $i = 0
foreach ($c in 'A', 'B', 'C', 'D', 'E', 'F', 'G') { $i++; $row += V "Sheet1!$($c)1" ($i * 10) }
$cases += @{ name = 'legacy-xls/implicit-intersection-row-range'; sheets = @('Sheet1', 'Sheet2'); formats = @('xls', 'xlsx')
    cells = $row + @((F 'Sheet2!A5' '=Sheet1!A1:G1'), (F 'Sheet2!D5' '=Sheet1!A1:G1'), (F 'Sheet2!G5' '=Sheet1!A1:G1'),
        (F 'Sheet2!H5' '=Sheet1!A1:G1'), (F 'Sheet2!D6' '=Sheet1!A1:G1*2'), (F 'Sheet2!D7' '=SUM(Sheet1!A1:G1)'),
        (F 'Sheet1!C3' '=A1:G1')) }
$col = @(); foreach ($r in 1..7) { $col += V "Sheet1!A$r" ($r * 100) }
$cases += @{ name = 'legacy-xls/implicit-intersection-column-range'; sheets = @('Sheet1', 'Sheet2'); formats = @('xls', 'xlsx')
    cells = $col + @((F 'Sheet2!C3' '=Sheet1!A1:A7'), (F 'Sheet2!C7' '=Sheet1!A1:A7'), (F 'Sheet2!C9' '=Sheet1!A1:A7'),
        (F 'Sheet2!D4' '=Sheet1!A1:A7+1'), (F 'Sheet2!E4' '=SUM(Sheet1!A1:A7)')) }

# ---------------------------------------------------------------- 7. false circular references
$cases += @{ name = 'circular/sumifs-unselected-cell-not-circular'; cells = @(
        (V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'y'), (F 'A1' '=SUMIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumif-unselected-cell-not-circular'; cells = @(
        (V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'y'), (F 'A1' '=SUMIF(B2:C2,"x",B1:C1)'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-selected-cell-truly-circular'; cells = @(
        (V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'x'), (F 'A1' '=SUMIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-selected-cell-iterate-on'; iterate = $true; maxIter = 100; maxChange = 0.001; cells = @(
        (V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'x'), (F 'A1' '=SUMIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1/2')) }
$cases += @{ name = 'circular/sumifs-later-criteria-range'; cells = @(
        (V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'y'), (V 'B3' 7), (V 'C3' 9),
        (F 'A1' '=SUMIFS(B3:C3,B2:C2,"x",B1:C1,">0")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-whole-column'; cells = @(
        (V 'B1' 1), (V 'B2' 2), (V 'B3' 3), (T 'C1' 'x'), (T 'C2' 'x'), (T 'C3' 'x'), (T 'C5' 'y'),
        (F 'D1' '=SUMIFS(B:B,C:C,"x")'), (F 'B5' '=D1*2')) }
$cases += @{ name = 'circular/index-control'; cells = @(
        (V 'B1' 5), (F 'A1' '=INDEX(B1:C1,1,1)'), (F 'C1' '=A1*2')) }

# ---------------------------------------------------------------- 8. deep dependency chains
# Each cell reads the row below, so calculation order runs against row order.
$cases += @{ name = 'dependency-chain/up-5000-rows'; list = $false
    cells = @((F 'A1:A4999' '=A2+1'), (V 'A5000' 1)) }
$cases += @{ name = 'dependency-chain/down-5000-rows-control'; list = $false
    cells = @((V 'A1' 1), (F 'A2:A5000' '=A1+1')) }
# Sheet1!A{r} reads Sheet2!A{r}, which reads Sheet1!A{r+1}: 4,000 links hopping between sheets.
$cases += @{ name = 'dependency-chain/up-2000-rows-two-sheet-hop'; sheets = @('Sheet1', 'Sheet2'); list = $false
    cells = @((F 'Sheet1!A1:A2000' '=Sheet2!A1+1'), (F 'Sheet2!A1:A1999' '=Sheet1!A2+1'), (V 'Sheet2!A2000' 1)) }

# ---------------------------------------------------------------- 9. RATE with pmt 0
$cases += @{ name = 'rate/zero-pmt-zero-fv'; cells = @(Column @(
            '=RATE(5,0,-100,0)', '=RATE(5,0,100,0)', '=RATE(10,0,-1000,0)', '=RATE(5,0,-100,0,0.1)',
            '=RATE(1,0,-100,0)', '=RATE(5,0,0,0)')) }
$cases += @{ name = 'rate/zero-pmt-tiny-fv'; cells = @(Column @(
            '=RATE(5,0,-100,1E-4)', '=RATE(5,0,-100,-1E-4)', '=RATE(5,0,-100,1E-10)')) }

# ---------------------------------------------------------------- 10. round-trip features
$pivotData = @((T 'A1' 'Category'), (T 'B1' 'Amount'),
    (T 'A2' 'a'), (V 'B2' 10), (T 'A3' 'b'), (V 'B3' 20), (T 'A4' 'a'), (V 'B4' 30),
    (T 'A5' 'c'), (V 'B5' 40), (T 'A6' 'b'), (V 'B6' 50), (T 'A7' 'a'), (V 'B7' 60))
$cases += @{ name = 'round-trip/pivot-row-field-sum'; list = $false; cells = $pivotData
    build = {
        param($wb, $xl)
        $ws = $wb.Worksheets.Item('Sheet1')
        $pc = $wb.PivotCaches().Create(1, 'Sheet1!R1C1:R7C2')        # xlDatabase
        $pt = $pc.CreatePivotTable($ws.Range('D3'), 'PivotTable1')
        $pt.PivotFields('Category').Orientation = 1                   # xlRowField
        [void]$pt.AddDataField($pt.PivotFields('Amount'), 'Sum of Amount', -4157)  # xlSum
        $ws.Range('H1').Formula = '=GETPIVOTDATA("Amount",$D$3,"Category","b")'
        $ws.Range('H2').Formula = '=SUMIF(A2:A7,"b",B2:B7)'
    } }
$cases += @{ name = 'round-trip/chart-column-clustered'; list = $false
    cells = @((T 'A1' 'Item'), (T 'B1' 'Value'), (T 'A2' 'p'), (V 'B2' 3), (T 'A3' 'q'), (V 'B3' 5),
        (T 'A4' 'r'), (V 'B4' 2), (T 'A5' 's'), (V 'B5' 7), (F 'B6' '=SUM(B2:B5)'))
    build = {
        param($wb, $xl)
        $ws = $wb.Worksheets.Item('Sheet1')
        $co = $ws.ChartObjects().Add(200, 10, 320, 200)
        $co.Chart.ChartType = 51                                       # xlColumnClustered
        $co.Chart.SetSourceData($ws.Range('A1:B5'))
    } }

# ================================================================ batch 2

# ---------------------------------------------------------------- 11. pivot sources outside the workbook
# pivot-external/source-data.xlsx is the source: a plain range plus a workbook-level name over
# it. The pivot cases copy it into the staging folder and open it from there, so the saved
# workbooks link to it by a bare relative name and resolve next to it in excel/pivot-external/.
$pivotRows = @((T 'A1' 'Region'), (T 'B1' 'Category'), (T 'C1' 'Amount'),
    (T 'A2' 'north'), (T 'B2' 'a'), (V 'C2' 11), (T 'A3' 'south'), (T 'B3' 'b'), (V 'C3' 23),
    (T 'A4' 'north'), (T 'B4' 'b'), (V 'C4' 35), (T 'A5' 'east'), (T 'B5' 'a'), (V 'C5' 47),
    (T 'A6' 'south'), (T 'B6' 'a'), (V 'C6' 59), (T 'A7' 'north'), (T 'B7' 'a'), (V 'C7' 61),
    (T 'A8' 'east'), (T 'B8' 'c'), (V 'C8' 72), (T 'A9' 'south'), (T 'B9' 'c'), (V 'C9' 84))
function Open-PivotSource($xl, $ctx) {
    $src = Join-Path $ctx.outRoot 'pivot-external\source-data.xlsx'
    if (-not (Test-Path -LiteralPath $src)) { throw 'pivot-external/source-data.xlsx is missing; generate it first' }
    $copy = Join-Path $ctx.stage 'source-data.xlsx'
    Copy-Item -LiteralPath $src -Destination $copy -Force
    return $xl.Workbooks.Open($copy)
}
function Add-Pivot($wb, $pc, [string]$at, [string]$name, [string[]]$rowFields) {
    $ws = $wb.Worksheets.Item('Sheet1')
    $pt = $pc.CreatePivotTable($ws.Range($at), $name)
    $pos = 0
    foreach ($f in $rowFields) { $pos++; $pf = $pt.PivotFields($f); $pf.Orientation = 1; $pf.Position = $pos }   # xlRowField
    [void]$pt.AddDataField($pt.PivotFields('Amount'), 'Sum of Amount', -4157)                                # xlSum
}
$cases += @{ name = 'pivot-external/source-data'; cells = $pivotRows
    build = { param($wb, $xl) [void]$wb.Names.Add('SourceRows', '=Sheet1!$A$1:$C$9') } }
$extFormulas = @((F 'H1' '=GETPIVOTDATA("Amount",$D$3,"Category","b")'), (F 'H2' '=GETPIVOTDATA("Amount",$D$3)'),
    (F 'H3' '=GETPIVOTDATA("Sum of Amount",$D$3,"Category","c")'))
$cases += @{ name = 'pivot-external/range-source'; cells = $extFormulas
    build = {
        param($wb, $xl, $ctx)
        $src = Open-PivotSource $xl $ctx
        $wb.Activate()
        $pc = $wb.PivotCaches().Create(1, '[source-data.xlsx]Sheet1!R1C1:R9C3')                     # xlDatabase
        Add-Pivot $wb $pc 'D3' 'PivotTable1' @('Category')
    } }
$cases += @{ name = 'pivot-external/defined-name-source'; cells = $extFormulas
    build = {
        param($wb, $xl, $ctx)
        $src = Open-PivotSource $xl $ctx
        $wb.Activate()
        $pc = $wb.PivotCaches().Create(1, 'source-data.xlsx!SourceRows')
        Add-Pivot $wb $pc 'D3' 'PivotTable1' @('Category')
    } }
$cases += @{ name = 'pivot-external/shared-cache'
    cells = $extFormulas + @((F 'H4' '=GETPIVOTDATA("Amount",$D$12,"Region","north")'), (F 'H5' '=H1+H4'))
    build = {
        param($wb, $xl, $ctx)
        $src = Open-PivotSource $xl $ctx
        $wb.Activate()
        $pc = $wb.PivotCaches().Create(1, '[source-data.xlsx]Sheet1!R1C1:R9C3')
        Add-Pivot $wb $pc 'D3' 'PivotTable1' @('Category')
        Add-Pivot $wb $pc 'D12' 'PivotTable2' @('Region')
    } }

# ---------------------------------------------------------------- 12. GETPIVOTDATA on an in-sheet pivot
# Pivot at E3 over A1:C9, rows Region then Category (compact layout, subtotals on).
$gpdBuild = {
    param($wb, $xl)
    $pc = $wb.PivotCaches().Create(1, 'Sheet1!R1C1:R9C3')
    Add-Pivot $wb $pc 'E3' 'PivotTable1' @('Region', 'Category')
}
$cases += @{ name = 'getpivotdata/field-item-forms'; build = $gpdBuild
    cells = $pivotRows + @((T 'A11' 'north'),
        (F 'I1' '=GETPIVOTDATA("Amount",$E$3)'),
        (F 'I2' '=GETPIVOTDATA("Amount",$E$3,"Region","north")'),
        (F 'I3' '=GETPIVOTDATA("Amount",$E$3,"Region","north","Category","a")'),
        (F 'I4' '=GETPIVOTDATA("Sum of Amount",$E$3,"Region","south")'),
        (F 'I5' '=GETPIVOTDATA("amount",$E$3,"region","EAST","category","C")'),
        (F 'I6' '=GETPIVOTDATA("Amount",$F$6,"Region",A11)'),
        (F 'I7' '=GETPIVOTDATA("Amount",$E$3,"Category","b")'),
        (F 'I8' '=GETPIVOTDATA("Amount",$E$3,"Region","west")'),
        (F 'I9' '=GETPIVOTDATA("Amount",$E$3,"Region","east","Category","b")'),
        (F 'I10' '=GETPIVOTDATA("Total",$E$3)')) }
# Control: the same pivot, and the same totals from SUMIFS over the source rows.
$cases += @{ name = 'getpivotdata/sumifs-control'; build = $gpdBuild
    cells = $pivotRows + @((F 'I1' '=SUM(C2:C9)'), (F 'I2' '=SUMIFS(C2:C9,A2:A9,"north")'),
        (F 'I3' '=SUMIFS(C2:C9,A2:A9,"north",B2:B9,"a")'), (F 'I4' '=SUMIFS(C2:C9,A2:A9,"south")')) }


# ---------------------------------------------------------------- 12b. GETPIVOTDATA shapes
# Ten more shapes, each probed in Excel 16.0 en-US: row and column fields, report filters, totals
# off, dates and numbers as items, renamed captions, hidden items, several data fields, a stale
# report (source edited, PivotTable not refreshed), and a tabular PivotTable on another sheet.
# Source: Region, Category, Amount, Date (serial) and Size over A1:E9. The formulas are in Q.
$gpdRows5 = @((T 'A1' 'Region'), (T 'B1' 'Category'), (T 'C1' 'Amount'), (T 'D1' 'Date'), (T 'E1' 'Size'),
    (T 'A2' 'north'), (T 'B2' 'a'), (V 'C2' 11), (V 'D2' 46037 'm/d/yyyy'), (V 'E2' 10),
    (T 'A3' 'south'), (T 'B3' 'b'), (V 'C3' 23), (V 'D3' 46037 'm/d/yyyy'), (V 'E3' 20),
    (T 'A4' 'north'), (T 'B4' 'b'), (V 'C4' 35), (V 'D4' 46054 'm/d/yyyy'), (V 'E4' 10),
    (T 'A5' 'east'), (T 'B5' 'a'), (V 'C5' 47), (V 'D5' 46054 'm/d/yyyy'), (V 'E5' 20),
    (T 'A6' 'south'), (T 'B6' 'a'), (V 'C6' 59), (V 'D6' 46037 'm/d/yyyy'), (V 'E6' 10),
    (T 'A7' 'north'), (T 'B7' 'a'), (V 'C7' 61), (V 'D7' 46091 'm/d/yyyy'), (V 'E7' 20),
    (T 'A8' 'east'), (T 'B8' 'c'), (V 'C8' 72), (V 'D8' 46091 'm/d/yyyy'), (V 'E8' 10),
    (T 'A9' 'south'), (T 'B9' 'c'), (V 'C9' 84), (V 'D9' 46054 'm/d/yyyy'), (V 'E9' 20))
function New-GpdPivot($wb, [string]$at, [string]$name, [string[]]$rowF, [string[]]$colF, [string[]]$pageF, [string]$sheet = 'Sheet1') {
    $pc = $wb.PivotCaches().Create(1, 'Sheet1!R1C1:R9C5')
    $pt = $pc.CreatePivotTable($wb.Worksheets.Item($sheet).Range($at), $name)
    $pos = 0; foreach ($f in $rowF) { $pos++; $pf = $pt.PivotFields($f); $pf.Orientation = 1; $pf.Position = $pos }   # xlRowField
    $pos = 0; foreach ($f in $colF) { $pos++; $pf = $pt.PivotFields($f); $pf.Orientation = 2; $pf.Position = $pos }   # xlColumnField
    foreach ($f in $pageF) { $pt.PivotFields($f).Orientation = 3 }                                                    # xlPageField
    return $pt
}
function Add-GpdSum($pt, [string]$caption = 'Sum of Amount', [int]$fn = -4157) {
    return $pt.AddDataField($pt.PivotFields('Amount'), $caption, $fn)
}
$cases += @{ name = 'getpivotdata/row-col-fields'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Category","a")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3,"Region","south","Category","c")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",$G$3,"Category","c","Region","east")'),
        (F 'Q6' '=GETPIVOTDATA("Amount",$G$3,"Region","east","Category","b")'),
        (F 'Q7' '=GETPIVOTDATA("Amount",$H$5,"Category","b")'),
        (F 'Q8' '=GETPIVOTDATA("Amount",$A$20)'),
        (F 'Q9' '=GETPIVOTDATA("Amount",$G$3,"Category","d")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region') @('Category') @(); [void](Add-GpdSum $pt)
    } }
$cases += @{ name = 'getpivotdata/report-filter'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$5)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$5,"Category","a")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$5,"Region","north")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$5,"Region","south")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",$G$5,"Region","north","Category","b")'),
        (F 'Q6' '=GETPIVOTDATA("Amount",$G$3,"Category","a")'),
        (F 'Q7' '=GETPIVOTDATA("Amount",$G$20)'),
        (F 'Q8' '=GETPIVOTDATA("Amount",$G$20,"Region","north")'),
        (F 'Q9' '=GETPIVOTDATA("Amount",$G$20,"Category","c")'),
        (F 'Q10' '=GETPIVOTDATA("Amount",$G$35)'),
        (F 'Q11' '=GETPIVOTDATA("Amount",$G$35,"Region","north")'),
        (F 'Q12' '=GETPIVOTDATA("Amount",$G$35,"Category","a")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G5' 'PT1' @('Category') @() @('Region'); [void](Add-GpdSum $pt)
        $pt.PivotFields('Region').CurrentPage = 'north'
        $p2 = New-GpdPivot $wb 'G20' 'PT2' @('Category') @() @('Region'); [void](Add-GpdSum $p2)
        $p3 = New-GpdPivot $wb 'G35' 'PT3' @('Category') @() @('Region'); [void](Add-GpdSum $p3)
        $pf = $p3.PivotFields('Region'); $pf.EnableMultiplePageItems = $true; $pf.PivotItems('east').Visible = $false
    } }
$cases += @{ name = 'getpivotdata/totals-off'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Region","north","Category","a")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3,"Region","south","Category","c")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region', 'Category') @() @(); [void](Add-GpdSum $pt)
        $pt.PivotFields('Region').Subtotals = [object[]](@($false) * 12)
        $pt.ColumnGrand = $false; $pt.RowGrand = $false
    } }
$cases += @{ name = 'getpivotdata/date-number-items'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3,"Date",DATE(2026,1,15))'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Date",46054)'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Date","3/10/2026")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3,"Date","2026-01-15")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",$G$3,"Date",DATE(2026,1,16))'),
        (F 'Q6' '=GETPIVOTDATA("Amount",$G$20,"Size",10)'),
        (F 'Q7' '=GETPIVOTDATA("Amount",$G$20,"Size","20")'),
        (F 'Q8' '=GETPIVOTDATA("Amount",$G$20,"Size",10.5)'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Date') @() @(); [void](Add-GpdSum $pt)
        # Undo automatic date grouping if Excel applied it.
        try { if ($pt.PivotFields().Count -gt 5) { [void]$pt.PivotFields('Date').DataRange.Cells.Item(1).Ungroup() } } catch { }
        $p2 = New-GpdPivot $wb 'G20' 'PT2' @('Size') @() @(); [void](Add-GpdSum $p2)
    } }
$cases += @{ name = 'getpivotdata/captions'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Total Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q3' '=GETPIVOTDATA("Sum of Amount",$G$3)'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3,"Area","south")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",$G$3,"Region","south")'),
        (F 'Q6' '=GETPIVOTDATA("Amount",$G$3,"Region","North Zone")'),
        (F 'Q7' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'),
        (F 'Q8' '=GETPIVOTDATA("Amount",$G$3,"Area","North Zone","Category","a")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region', 'Category') @() @(); [void](Add-GpdSum $pt 'Total Amount')
        $pf = $pt.PivotFields('Region')
        $pf.PivotItems('north').Caption = 'North Zone'
        $pf.Caption = 'Area'
    } }
$cases += @{ name = 'getpivotdata/hidden-items'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Region","south")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Region","east")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3,"Region","east","Category","c")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",$G$3,"Region","north","Category","b")'),
        (F 'Q6' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region', 'Category') @() @(); [void](Add-GpdSum $pt)
        $pt.PivotFields('Region').PivotItems('south').Visible = $false
        $pt.PivotFields('Category').PivotItems('c').Visible = $false
    } }
$cases += @{ name = 'getpivotdata/multi-data-fields'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Sum of Amount",$G$3,"Region","north")'),
        (F 'Q2' '=GETPIVOTDATA("Count of Amount",$G$3,"Region","north")'),
        (F 'Q3' '=GETPIVOTDATA("Average of Amount",$G$3,"Region","south")'),
        (F 'Q4' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q5' '=GETPIVOTDATA("Pct of Amount",$G$3,"Region","east")'),
        (F 'Q6' '=GETPIVOTDATA("Count of Amount",$G$3)'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region') @() @()
        [void](Add-GpdSum $pt 'Sum of Amount')
        [void](Add-GpdSum $pt 'Count of Amount' -4112)
        [void](Add-GpdSum $pt 'Average of Amount' -4106)
        $df = Add-GpdSum $pt 'Pct of Amount'; $df.Calculation = 8   # xlPercentOfTotal
    } }
$cases += @{ name = 'getpivotdata/stale-report'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Region","south")'),
        (F 'Q4' '=SUMIFS(C2:C9,A2:A9,"north")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region') @() @(); [void](Add-GpdSum $pt)
        $pt.ManualUpdate = $true
        $wb.Worksheets.Item('Sheet1').Range('C2').Value2 = 1011   # north/a 11 -> 1011, PivotTable not refreshed
    } }
$cases += @{ name = 'getpivotdata/stale-report-sorted'; cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",$G$3)'),
        (F 'Q2' '=GETPIVOTDATA("Amount",$G$3,"Region","north")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",$G$3,"Region","south")'),
        (F 'Q4' '=SUMIFS(C2:C9,A2:A9,"north")'))
    build = {
        param($wb, $xl)
        $ws = $wb.Worksheets.Item('Sheet1')
        # Source rows sorted by Region, so first-appearance order equals Excel's ascending order:
        # separates a stale report from an item-order difference.
        $sorted = @(@('east', 'a', 47), @('east', 'c', 72), @('north', 'a', 11), @('north', 'b', 35),
            @('north', 'a', 61), @('south', 'b', 23), @('south', 'a', 59), @('south', 'c', 84))
        for ($r = 0; $r -lt 8; $r++) { for ($c = 0; $c -lt 3; $c++) { $v = $sorted[$r][$c]; if ($c -eq 2) { $ws.Cells.Item($r + 2, $c + 1).Value2 = [double]$v } else { $ws.Cells.Item($r + 2, $c + 1).Value2 = [string]$v } } }
        $pt = New-GpdPivot $wb 'G3' 'PT1' @('Region') @() @(); [void](Add-GpdSum $pt)
        $ws.Range('C4').Value2 = 1011   # north/a 11 -> 1011, not refreshed
    } }
$cases += @{ name = 'getpivotdata/tabular-other-sheet'; sheets = @('Sheet1', 'Pivot'); cells = $gpdRows5 + @((F 'Q1' '=GETPIVOTDATA("Amount",Pivot!$A$3,"Region","north","Category","a")'),
        (F 'Q2' '=GETPIVOTDATA("Amount",Pivot!$A$3,"Region","east")'),
        (F 'Q3' '=GETPIVOTDATA("Amount",Pivot!$A$3)'),
        (F 'Q4' '=GETPIVOTDATA("Amount",Pivot!$A$3,"Category","a")'),
        (F 'Q5' '=GETPIVOTDATA("Amount",Pivot!$A$1:$H$40,"Region","south")'))
    build = {
        param($wb, $xl)
        $pt = New-GpdPivot $wb 'A3' 'PT1' @('Region', 'Category') @() @() 'Pivot'; [void](Add-GpdSum $pt)
        $pt.RowAxisLayout(1)   # xlTabularRow
    } }

# ---------------------------------------------------------------- 13. more date text
# Probed in Excel en-US first; each file holds what Excel actually does with the form.
# Day in the current year: a number 1-31 after a month name with "-" or "/" is a day, and long
# month prefixes ("Janu", "Febr") are accepted. Cached serials are right only during 2026.
$cases += @{ name = 'date-text/month-name-day-this-year-y2026'; cells = @(Column @(
            '=DATEVALUE("Jan-21")', '=DATEVALUE("Jan/21")', '=DATEVALUE("Jan-1")', '=DATEVALUE("Jan/1")',
            '=DATEVALUE("Jan-30")', '="Jan-21"+0', '=DATEVALUE("1Janu")', '=DATEVALUE("Janu 1")',
            '=DATEVALUE("1Febr")', '=DATEVALUE("01Jan")', '=DATEVALUE("Sep 1")')) }
# The same number-as-year reading with separators and with "+0" coercion (the space forms are
# in month-name-number-as-year.xlsx).
$cases += @{ name = 'date-text/month-name-number-as-year-separators'; cells = @(Column @(
            '=DATEVALUE("Jan-32")', '=DATEVALUE("Jan/32")', '=DATEVALUE("Jan-0")', '=DATEVALUE("Feb-30")',
            '="Jan 32"+0', '="Feb 30"+0', '="Jan 0"+0', '=YEAR("Jan 32")')) }
# Two-digit years: 29 is 2029, 30 is 1930 (Windows' 1930-2029 window).
$cases += @{ name = 'date-text/two-digit-year-window'; cells = @(Column @(
            '=DATEVALUE("1Jan29")', '=DATEVALUE("1Jan30")', '=DATEVALUE("1-Jan-29")', '=DATEVALUE("1-Jan-30")',
            '=DATEVALUE("1 Jan 29")', '=DATEVALUE("1 Jan 30")', '=YEAR("1Jan29")', '=YEAR("1Jan30")')) }
# Month prefixes longer than three letters, with an explicit year.
$cases += @{ name = 'date-text/month-name-long-prefix'; cells = @(Column @(
            '=DATEVALUE("1 Sept 2021")', '=DATEVALUE("1Sept2021")', '=DATEVALUE("Sept 2021")', '=DATEVALUE("1Janu2021")',
            '=DATEVALUE("Janu 2021")', '=DATEVALUE("Septe 2021")', '=DATEVALUE("1 Septem 2021")', '=DATEVALUE("Decemb 2021")',
            '=DATEVALUE("Janux 2021")', '=DATEVALUE("Sept")')) }
# Forms that stay #VALUE!: dots, a comma after the day, year first, a year after "Mon D"
# without a comma, a three-digit day, a two-letter month.
$cases += @{ name = 'date-text/month-name-punctuation-value-errors'; cells = @(Column @(
            '=DATEVALUE("4.Mar.2026")', '=DATEVALUE("Jan.21")', '=DATEVALUE("1Jan, 2021")', '=DATEVALUE("4 Mar, 2026")',
            '=DATEVALUE("2021Jan")', '=DATEVALUE("Jan 15 2021")', '=DATEVALUE("001Jan")', '=DATEVALUE("Ma")',
            '=DATEVALUE("Sept 1 2021")', '=DATEVALUE("Mar.4")', '="Jan 15 2021"+0')) }
# Control: the punctuated forms Excel does accept, all with an explicit four-digit year.
$cases += @{ name = 'date-text/month-name-punctuation-control'; cells = @(Column @(
            '=DATEVALUE("Mar 4, 2026")', '=DATEVALUE("Jan 15, 2021")', '=DATEVALUE("4-Mar-2026")', '=DATEVALUE("4/Mar/2026")',
            '=DATEVALUE("Jan-2021")', '=DATEVALUE("Jan/2021")')) }

# ---------------------------------------------------------------- 14. precision as displayed, by format family
# Each source is a formula in a formatted cell; the General cell next to it reads the stored
# (rounded) value. Excel 365 rounds to what the format shows for every number format probed
# except dates, times, General and text.
function PadPair([string]$row, [string]$f, [string]$fmt) { (F "A$row" $f $fmt), (F "B$row" "=A$row*3") }
$padFamilies = [ordered]@{
    'percent'           = @(@('=1/3', '0%'), @('=0.4+0.003', '0.0%'), @('=-2/3', '0.00%'))
    'currency'          = @(@('=10/3', '$#,##0.00'), @('=-10/3', '$#,##0.00;($#,##0.00)'), @('=1234.5678', '"$"#,##0.0'))
    'thousands-scaling' = @(@('=1234567.891', '0.0,,'), @('=12345.678', '#,##0,'), @('=9876.54', '0.00,'))
    'fractions'         = @(@('=2+1/3', '# ?/?'), @('=0.3+0', '# ?/?'), @('=0.3+0', '?/8'), @('=-10/3', '?/8'), @('=45000.123456', '# ?/?'))
    'scientific'        = @(@('=12345.678', '##0.0E+0'), @('=1/7', '0.00E+00'), @('=-98765.4', '0.0E+0'))
    'multi-section'     = @(@('=10/3', '0.00;(0.0)'), @('=-10/3', '0.00;(0.0)'), @('=-10/7', '0;-0.000'))
    'literal-text'      = @(@('=10/3', '0.0" u"'), @('=20/3', '"x"0.00'), @('=10/3', '0.0"%"'))
    # Not a control, despite the condition: Excel rounds by the first section's decimals
    # (2+1/3 shows "2" but is stored as 2.33; -10/3 is stored as -3).
    'conditional-sections' = @(@('=1/3', '[<1]0.00;0'), @('=2+1/3', '[<1]0.00;0'), @('=-10/3', '[<1]0.00;0'), @('=12345.678', '[<1]0.00;0'))
    # Controls: kept at full value.
    'general-control'   = @(@('=1/3', 'General'), @('=10/3', 'General'), @('=-2/7', 'General'))
    'date-time-control' = @(@('=45000.123456', 'm/d/yyyy'), @('=1/3+0.0001', 'h:mm'), @('=45000.123456', 'm/d/yyyy h:mm'), @('=2+1/3', '[h]:mm'))
}
foreach ($k in $padFamilies.Keys) {
    $cells = @(); $r = 0
    foreach ($p in $padFamilies[$k]) { $r++; $cells += PadPair "$r" $p[0] $p[1] }
    $cases += @{ name = "precision-displayed/pad-on-$k"; cells = $cells; pad = $true }
}
# Text format "@" set after the formulas (a formula typed into a text cell would be text).
$cases += @{ name = 'precision-displayed/pad-on-text-format-control'; pad = $true
    cells = @((F 'A1' '=10/3'), (F 'B1' '=A1*3'), (F 'A2' '=-2/7'), (F 'B2' '=A2*7'))
    build = { param($wb, $xl) $wb.Worksheets.Item(1).Range('A1:A2').NumberFormat = '@' } }
# A formatted cell reading a General one rounds; a General cell reading a formatted one reads
# the rounded value; a General cell reading a General one stays exact. Excel copies A2's format
# into B2 when =A2 is entered, so B2 is set back to General afterwards.
$cases += @{ name = 'precision-displayed/pad-on-cross-cell-read'; pad = $true; cells = @(
        (F 'A1' '=10/3'), (F 'B1' '=A1' '0.0'), (F 'C1' '=B1*3'), (F 'D1' '=A1*3'),
        (F 'A2' '=20/3' '0.0'), (F 'B2' '=A2'), (F 'C2' '=B2*3'))
    build = { param($wb, $xl) $wb.Worksheets.Item(1).Range('B2').NumberFormat = 'General' } }
# One spilled array landing in cells formatted 0, 0.0 and 0.00: each element is rounded by
# the format of the cell it lands in.
$cases += @{ name = 'precision-displayed/pad-on-spill-mixed-formats'; pad = $true; cells = @(
        @{ a = 'A1'; n = '0' }, @{ a = 'A2'; n = '0.0' }, @{ a = 'A3'; n = '0.00' },
        @{ a = 'A1'; f2 = '=SEQUENCE(3,1,1/3,1/3)' },
        (F 'B1:B3' '=A1*3'), (F 'C1' '=SUM(A1:A3)'))
    # Excel copies A1's "0" into C1 when the SUM is entered; C1 is meant to be General.
    build = { param($wb, $xl) $wb.Worksheets.Item(1).Range('C1').NumberFormat = 'General' } }

# ---------------------------------------------------------------- 15. more false circular references
# Iteration off throughout. C1 depends on A1; whether A1's formula really reads C1 decides.
$circBase = @((V 'B1' 5), (T 'B2' 'x'), (T 'C2' 'y'), (V 'B3' 7), (V 'C3' 9))
# The twin of sumifs-later-criteria-range with the two criteria pairs swapped.
$cases += @{ name = 'circular/sumifs-first-criteria-range'; cells = $circBase + @(
        (F 'A1' '=SUMIFS(B3:C3,B1:C1,">0",B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-size-mismatch-value-error'; cells = $circBase + @(
        (F 'A1' '=SUMIFS(B1:C1,B2:D2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-array-criterion-unselected'; cells = $circBase + @(
        (F 'A1' '=SUM(SUMIFS(B1:C1,B2:C2,{"x","z"}))'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/sumifs-array-criterion-selected'; cells = $circBase + @(
        (F 'A1' '=SUM(SUMIFS(B1:C1,B2:C2,{"x","y"}))'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/averageifs-unselected-cell'; cells = $circBase + @(
        (F 'A1' '=AVERAGEIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/maxifs-unselected-cell'; cells = $circBase + @(
        (F 'A1' '=MAXIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/minifs-unselected-cell'; cells = $circBase + @(
        (F 'A1' '=MINIFS(B1:C1,B2:C2,"x")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/countifs-later-criteria-range'; cells = $circBase + @(
        (F 'A1' '=COUNTIFS(B2:C2,"x",B1:C1,">0")'), (F 'C1' '=A1*2')) }
$cases += @{ name = 'circular/countifs-first-criteria-range'; cells = $circBase + @(
        (F 'A1' '=COUNTIFS(B1:C1,">0",B2:C2,"x")'), (F 'C1' '=A1*2')) }

# ---------------------------------------------------------------- 16. more legacy formulas in .xls
$cases += @{ name = 'legacy-xls/cse-single-cell-array'; formats = @('xls', 'xlsx'); cells = @(
        (V 'A1' 1), (V 'A2' 2), (V 'A3' 3), (V 'B1' 4), (V 'B2' 5), (V 'B3' 6),
        @{ a = 'D1'; fa = '=SUM(A1:A3*B1:B3)' }, @{ a = 'D2'; fa = '=MAX(A1:A3-B1:B3)' },
        (F 'D3' '=SUMPRODUCT(A1:A3,B1:B3)')) }
# A formula filled down a range is stored as a shared formula (SHRFMLA in .xls).
$cases += @{ name = 'legacy-xls/shared-formula-fill-down'; formats = @('xls', 'xlsx'); cells = @(
        (V 'A1' 3), (V 'A2' 1), (V 'A3' 4), (V 'A4' 1), (V 'A5' 5), (V 'A6' 9),
        (F 'B1:B6' '=A1*2+1'), (F 'C1:C6' '=B1+$A$1'), (F 'D1' '=SUM(C1:C6)')) }
# A range inside an operator inside SUM, entered as a legacy formula: implicitly intersected
# (the row's cell, or #VALUE! outside the range); the CSE copy sums the whole array.
$cases += @{ name = 'legacy-xls/implicit-intersection-inside-sum'; formats = @('xls', 'xlsx'); cells = @(
        (V 'A1' 12), (V 'A2' 24), (V 'A3' 36),
        (F 'C2' '=MAX(SUM(A1:A3/12),0)'), (F 'C5' '=MAX(SUM(A1:A3/12),0)'),
        @{ a = 'C1'; fa = '=MAX(SUM(A1:A3/12),0)' }, (F 'C3' '=SUM(A1:A3)/12')) }

# ---------------------------------------------------------------- 17. round-trip: document properties
# Saved with personal information removed, so Author and Last Author are empty, but
# docProps/core.xml still carries the created and modified dates, which a save must keep.
$cases += @{ name = 'round-trip/core-properties-blank-author'; cells = @((V 'A1' 2), (F 'A2' '=A1*21')) }


# ---------------------------------------------------------------- 17. Analysis ToolPak value arguments
# EOMONTH, WORKDAY, COUPDAYS, BIN2DEC and the other Analysis ToolPak functions give #VALUE! for a
# reference to more than one cell in a value argument (no implicit intersection, no spill), while
# array constants and computed arrays lift. Range arguments (holidays, cash flows) take ranges.
$atpData = @((V 'A1' 45000), (V 'A2' 45031), (V 'A3' 45061), (V 'B1' 1), (V 'B2' 2), (V 'B3' 3), (V 'C2' 7),
    (V 'F1' 1), (V 'F2' 2), (V 'F3' 11))
function Column2([string[]]$formulas, [string]$col, [int]$step = 1) {
    $i = 0
    foreach ($f in $formulas) { $i++; @{ a = "$col$(1 + ($i - 1) * $step)"; f2 = $f } }
}
$cases += @{ name = 'atp-array-args/date-functions'; cells = $atpData + @(Column2 @(
            '=EOMONTH(A1:A3,1)', '=EOMONTH(A1,B1:B3)', '=EDATE(A1:A3,1)', '=EDATE(A1,B1:B3)', '=WORKDAY(A1:A3,1)',
            '=WORKDAY(A1,B1:B3)', '=WORKDAY.INTL(A1:A3,1)', '=NETWORKDAYS(A1:A3,A3)', '=NETWORKDAYS.INTL(A1,A1:A3)',
            '=YEARFRAC(A1:A3,A3)', '=YEARFRAC(A1,A3,B1:B2)', '=WEEKNUM(A1:A3)', '=WEEKNUM(A1,B1:B3)',
            '=SUMPRODUCT(EOMONTH(A1:A3,0))', '=EOMONTH(A1:A1,0)', '=WORKDAY(A1,1,A1:A3)', '=NETWORKDAYS(A1,A3,A1:A3)') 'H') }
$cases += @{ name = 'atp-array-args/financial-functions'; cells = $atpData + @(Column2 @(
            '=COUPDAYS(A1:A3,46000,2,0)', '=COUPNCD(A1,46000,B1:B2,0)', '=DISC(A1:A3,46000,98,100,0)', '=DISC(A1,46000,98,100,B1:B3)',
            '=PRICE(A1:A3,46000,0.05,0.06,100,2,0)', '=YIELD(A1,46000,0.05,98,100,B1:B2,0)', '=ACCRINT(44800,45100,A1:A3,0.05,1000,2,0)',
            '=TBILLPRICE(A1:A3,45100,0.05)', '=EFFECT(B1:B3,4)', '=CUMIPMT(0.01,12,1000,B1:B3,12,0)', '=XNPV(B1:B2,B1:B3,A1:A3)',
            '=XNPV(0.1,B1:B3,A1:A3)', '=FVSCHEDULE(B1:B2,{0.1,0.2})', '=FVSCHEDULE(100,B1:B3)', '=COUPDAYS(A1:A1,46000,2,0)') 'H') }
$cases += @{ name = 'atp-array-args/math-engineering-functions'; cells = $atpData + @(Column2 @(
            '=ISEVEN(B1:B3)', '=ISODD(B1:B3)', '=QUOTIENT(C1:C3,2)', '=MROUND(B1:B3,2)', '=SQRTPI(B1:B3)', '=SERIESSUM(B1:B2,1,1,B1:B3)',
            '=SERIESSUM(2,1,1,B1:B3)', '=BIN2DEC(B1:B3)', '=DEC2HEX(B1:B3)', '=DEC2BIN(5,B1:B3)', '=CONVERT(B1:B3,"m","ft")',
            '=DELTA(B1:B3,1)', '=ERF(B1:B3)', '=BESSELJ(B1:B3,1)', '=COMPLEX(B1:B3,2)', '=IMABS(B1:B3)', '=IMSUM(B1:B3)') 'H') }
# Spilling: constants and computed arrays lift; a rejected reference adds no shape.
$cases += @{ name = 'atp-array-args/lifted-arrays'; cells = $atpData + @(Column2 @(
            '=EOMONTH(A1:A3+0,0)', '=EOMONTH(A1,{1,2})', '=EOMONTH(A1,{1;2})', '=EOMONTH(A1:A3,{1;2})', '=EOMONTH(TRANSPOSE(A1:A3),0)',
            '=MAP(A1:A3,LAMBDA(x,EOMONTH(x,0)))', '=NETWORKDAYS.INTL(A1,A3+30,F1:F3)', '=ISOWEEKNUM(A1:A3)', '=BITAND(B1:B3,1)',
            '=COUPDAYS(A1:A3+0,46000,2,0)') 'H' 4) }
# However the reference is formed. Formulas in row 2, inside A1:A3's rows.
$cases += @{ name = 'atp-array-args/reference-forms'; cells = $atpData + @(
        @{ a = 'H2'; f2 = '=EOMONTH(INDEX(A1:A3,0,1),0)' }, @{ a = 'I2'; f2 = '=EOMONTH(OFFSET(A1,0,0,3,1),0)' },
        @{ a = 'J2'; f2 = '=EOMONTH(IF(TRUE,A1:A3),0)' }, @{ a = 'K2'; f2 = '=EOMONTH(CHOOSE(1,A1:A3),0)' },
        @{ a = 'L2'; f2 = '=LET(x,A1:A3,EOMONTH(x,0))' }, @{ a = 'M2'; f2 = '=EOMONTH((A1,A2),0)' },
        @{ a = 'N2'; f2 = '=EOMONTH(A:A,0)' }, @{ a = 'O2'; f2 = '=EOMONTH(A1:INDEX(A1:A3,3),0)' },
        @{ a = 'P2'; f2 = '=EOMONTH(INDEX(A1:A3,2),0)' }, @{ a = 'Q2'; f2 = '=EOMONTH(OFFSET(A1,0,0,1,1),0)' },
        @{ a = 'R2'; f2 = '=EOMONTH(A1:A3 A2:C2,0)' }, @{ a = 'S2'; f2 = '=EOMONTH(@A1:A3,0)' })
    build = { param($wb, $xl) [void]$wb.Names.Add('Dates', '=Sheet1!$A$1:$A$3'); $wb.Worksheets.Item(1).Range('T2').Formula2 = '=EOMONTH(Dates,0)' } }
# Legacy formulas (Range.Formula): no implicit intersection for value arguments, unlike YEAR.
$cases += @{ name = 'atp-array-args/legacy-formulas'; formats = @('xlsx'); cells = $atpData + @(
        (F 'H2' '=EOMONTH(A1:A3,1)'), (F 'I2' '=WORKDAY(A1:A3,1)'), (F 'J2' '=COUPDAYS(A1:A3,46000,2,0)'), (F 'K2' '=BIN2DEC(B1:B3)'),
        (F 'L2' '=YEAR(A1:A3)'), (F 'M2' '=NETWORKDAYS.INTL(A1,A3,F1:F3)'), (F 'N2' '=WORKDAY(A1,1,A1:A3)')) }

$cases
