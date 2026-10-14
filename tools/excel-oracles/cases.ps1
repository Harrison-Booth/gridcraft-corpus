# The case list for generate.ps1. Dot-sourced; returns an array of case hashtables.
#
# A case:
#   name      path under excel/ without extension, e.g. 'near-zero/sub-reordered-parts-small'
#   cells     list of cell specs (helpers below):
#               F 'A1' '=A2+1' ['0.00']   formula (Range.Formula: legacy, implicit intersection)
#               V 'A1' 1.5 ['0.00']       number constant
#               T 'A1' 'x'                text constant (cell formatted as text first)
#             An address may carry a sheet ('Sheet2!A1') and may be a range
#             ('A1:A4999'): a formula with relative references then fills the range.
#   sheets    sheet names, default @('Sheet1')
#   formats   @('xlsx'), @('xls') or both; default @('xlsx')
#   iterate   iterative calculation on (default $false); maxIter / maxChange go with it
#   pad       precision as displayed (default $false)
#   build     optional scriptblock { param($wb, $xl) ... } run after the cells, before saving
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

$cases
