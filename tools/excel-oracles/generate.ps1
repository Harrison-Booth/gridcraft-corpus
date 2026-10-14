# Generates the excel/ oracle workbooks from cases.ps1 by driving Microsoft Excel through COM.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-oracles/generate.ps1 [regex] [-Force]
#
# The regex is matched against the case path (e.g. 'near-zero/sub-reordered-parts-small.xlsx'),
# so '^date-text/' or 'xls$' select a subset. With no regex every case is selected.
# Existing files are left alone (printed as 'skip') unless -Force is given, so a re-run doesn't
# churn SHA256SUMS: Excel never writes the same bytes twice (timestamps, document ids).
#
# One hidden Excel instance is used for the run, and no Excel setting is changed for good.
# Every workbook is saved with Workbook.RemovePersonalInformation on (File > Info > Inspect
# Workbook > "remove personal information on save"). Without it Excel 365 writes the signed-in
# Office identity into docProps/core.xml lastModifiedBy and the .xls WRITEACCESS record, whatever
# Application.UserName says, and the save folder into xl/workbook.xml (x15ac:absPath). With it,
# Author, Last Author and Company are saved empty and absPath is left out. Files are still saved
# through a neutral staging folder in case a future Excel writes the folder anyway.
# verify-metadata.py then checks every file.
param(
    [string]$Pattern = '.',
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$OutRoot = Join-Path $Root 'excel'
$Stage = Join-Path $env:SystemDrive 'gc-corpus-stage'
$FileFormat = @{ xlsx = 51; xls = 56 }   # xlOpenXMLWorkbook, xlExcel8

Add-Type -Namespace GcCorpus -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);'

$all = . (Join-Path $PSScriptRoot 'cases.ps1')

function Get-Range($wb, [string]$addr) {
    # The leading comma stops PowerShell from unrolling a multi-cell Range into its cells.
    if ($addr -match '^(.+)!(.+)$') { return , $wb.Worksheets.Item($Matches[1]).Range($Matches[2]) }
    return , $wb.Worksheets.Item(1).Range($addr)
}

function Write-Listing($wb, $cells) {
    # One text cell per formula cell, "C1: =A1-B1", in the second column right of the used range.
    # The text is read back from Excel after the build step, so it shows the formula as saved
    # (e.g. with #REF! after a sheet was deleted). It never starts with "=", so no reader takes
    # it for a formula.
    $bySheet = [ordered]@{}
    foreach ($c in $cells) {
        if (-not ($c.f -or $c.f2 -or $c.fa)) { continue }
        $sheet = $wb.Worksheets.Item(1).Name; $a = $c.a
        if ($a -match '^(.+)!(.+)$') { $sheet = $Matches[1]; $a = $Matches[2] }
        if (-not $bySheet.Contains($sheet)) { $bySheet[$sheet] = @() }
        $bySheet[$sheet] += $a
    }
    foreach ($sheet in $bySheet.Keys) {
        $ws = $null
        try { $ws = $wb.Worksheets.Item($sheet) } catch { continue }   # deleted by the build step
        $used = $ws.UsedRange
        $col = $used.Column + $used.Columns.Count + 1
        $r = 0
        foreach ($a in $bySheet[$sheet]) {
            $r++
            $first = $ws.Range($a).Cells.Item(1, 1)
            $cell = $ws.Cells.Item($r, $col)
            $cell.NumberFormat = '@'
            $cell.Value2 = "$($first.Address($false, $false)): $($first.Formula)"
        }
    }
}

function New-CaseWorkbook($xl, $case, [string]$fmt, [string]$stagePath) {
    $sheets = @(if ($case.sheets) { $case.sheets } else { 'Sheet1' })
    $xl.SheetsInNewWorkbook = $sheets.Count
    $wb = $xl.Workbooks.Add()
    try {
        for ($i = 0; $i -lt $sheets.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = $sheets[$i] }
        # Application-wide calculation state: set for every case so nothing carries over.
        $xl.Calculation = -4105                                 # xlCalculationAutomatic
        $xl.Iteration = [bool]$case.iterate
        $xl.MaxIterations = if ($case.maxIter) { [int]$case.maxIter } else { 100 }
        $xl.MaxChange = if ($case.maxChange) { [double]$case.maxChange } else { 0.001 }
        $wb.PrecisionAsDisplayed = $false
        foreach ($c in $case.cells) {
            $rg = Get-Range $wb $c.a
            if ($c.n) { $rg.NumberFormat = $c.n }
            if ($c.ContainsKey('t')) { $rg.NumberFormat = '@'; $rg.Value2 = [string]$c.t }
            elseif ($c.ContainsKey('v')) { $rg.Value2 = [double]$c.v }
            elseif ($c.ContainsKey('f2')) { $rg.Formula2 = $c.f2 }      # dynamic-array (spilling) formula
            elseif ($c.ContainsKey('fa')) { $rg.FormulaArray = $c.fa }  # Ctrl+Shift+Enter array formula
            elseif ($c.ContainsKey('f')) { $rg.Formula = $c.f }
        }
        # The build step also gets the staging and output folders, for cases that open a second
        # workbook (pivot-external/ opens its source from the staging folder, so the saved link is
        # relative).
        if ($case.build) { & $case.build $wb $xl @{ stage = $Stage; outRoot = $OutRoot } }
        if ($case.list -ne $false) { Write-Listing $wb $case.cells }
        if ($case.pad) { $wb.PrecisionAsDisplayed = $true }
        $xl.CalculateFull()
        $wb.RemovePersonalInformation = $true
        $wb.Worksheets.Item(1).Activate()
        [void]$wb.Worksheets.Item(1).Range('A1').Select()
        if (Test-Path -LiteralPath $stagePath) { Remove-Item -LiteralPath $stagePath -Force }
        $wb.SaveAs($stagePath, $FileFormat[$fmt])
    }
    finally {
        $wb.Close($false)
        # Close anything the build step opened (e.g. a pivot source workbook), unsaved.
        while ($xl.Workbooks.Count -gt 0) { $xl.Workbooks.Item(1).Close($false) }
    }
}

$jobs = @()
foreach ($case in $all) {
    $formats = @(if ($case.formats) { $case.formats } else { 'xlsx' })
    foreach ($fmt in $formats) {
        $rel = "$($case.name).$fmt"
        if ($rel -match $Pattern) { $jobs += , @($case, $fmt, $rel) }
    }
}
if ($jobs.Count -eq 0) { Write-Host "no case matches '$Pattern'"; exit 1 }

$failed = 0
$xl = $null; $xlPid = 0; $userName = $null
New-Item -ItemType Directory -Force -Path $Stage | Out-Null
try {
    $xl = New-Object -ComObject Excel.Application
    [void][GcCorpus.Win32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$xlPid)
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents = $false
    $xl.AutomationSecurity = 3
    $userName = $xl.UserName
    $sheetsInNew = $xl.SheetsInNewWorkbook
    Write-Host "Excel $($xl.Version) build $($xl.Build), country $($xl.International(1)), decimal '$($xl.International(3))', date order $($xl.International(32))"
    foreach ($job in $jobs) {
        $case, $fmt, $rel = $job
        $dest = Join-Path $OutRoot ($rel -replace '/', '\')
        if ((Test-Path -LiteralPath $dest) -and -not $Force) { Write-Host "skip  $rel (exists; -Force regenerates)"; continue }
        try {
            $stagePath = Join-Path $Stage ([IO.Path]::GetFileName($dest))
            New-CaseWorkbook $xl $case $fmt $stagePath
            New-Item -ItemType Directory -Force -Path (Split-Path $dest) | Out-Null
            Move-Item -LiteralPath $stagePath -Destination $dest -Force
            Write-Host "ok    $rel"
        }
        catch {
            $failed++
            Write-Host "FAIL  $rel  $(($_.Exception.Message -replace '\s+', ' ').Trim())"
            if ($env:GC_DEBUG) { Write-Host $_.InvocationInfo.PositionMessage }
        }
    }
}
finally {
    if ($xl) {
        try { $xl.SheetsInNewWorkbook = $sheetsInNew; $xl.Iteration = $false } catch { }
        try { $xl.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # Only ever stop the Excel this script started.
    if ($xlPid -ne 0) { Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $Stage -Recurse -Force -ErrorAction SilentlyContinue
}

# Metadata check over the whole folder; the original Excel user name is passed so it's never
# written into the repository.
$py = if ($env:GC_PYTHON) { $env:GC_PYTHON } else { 'python' }
$forbid = @()
if ($userName) { $forbid = @('--forbid', $userName) }
& $py -I (Join-Path $PSScriptRoot 'verify-metadata.py') $OutRoot @forbid
if ($LASTEXITCODE -ne 0) { $failed++ }
if ($failed -gt 0) { exit 1 }
