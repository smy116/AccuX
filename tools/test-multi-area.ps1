param([string]$Configuration = 'Debug')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dir = Join-Path $root "src\AccuX.AddIn\bin\$Configuration\net48"
foreach ($name in @('Microsoft.Office.Interop.Excel', 'office', 'AccuX.Core', 'AccuX.Host', 'AccuX.Modules.BasicFinance')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path $dir "$name.dll"))
}
function Assert-Equal($expected, $actual, $label) {
    if ($expected -ne $actual) { throw "$label : expected=$expected actual=$actual" }
}
$app = $null
$book = $null
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $book = $app.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $hostAdapter = [AccuX.Host.ExcelRangeOperationHost]::new($app, $null, $null)
    $pipeline = [AccuX.Core.Operations.RangeOperationPipeline]::new($hostAdapter, $null, $null)
    $mark = [AccuX.Host.Features.ExcelCellMarkHost]::new($app)
    $sheet.Range('A1:D4').Value2 = 1.234
    $sheet.Range('H1').Value2 = 10.456
    $sheet.Range('F2').Formula = '=H1+1'
    $sheet.Range('A1:B2,B2:C3,F2').Select() | Out-Null
    $target = $hostAdapter.CaptureTarget()
    Assert-Equal 8 $target.CellCount 'Overlap deduplication'
    Assert-Equal 8 $hostAdapter.Read($target).Cells.Count 'Read count'
    $context = [AccuX.Core.Operations.OperationContext]::new('test.round', 'test', $target, $null)
    $result = $pipeline.Execute([AccuX.Modules.BasicFinance.Rounding.RoundingTransform]::new(1), $context, $null)
    Assert-Equal $true $result.Success "Pipeline: $($result.Message)"
    Assert-Equal 8 $result.Stats.ProcessedCells 'Processed count'
    Assert-Equal 1.2 $sheet.Range('B2').Value2 'Overlap value'
    Assert-Equal 1.234 $sheet.Range('D4').Value2 'Unselected value'
    Assert-Equal '=ROUND(H1+1,1)' $sheet.Range('F2').Formula 'Formula stays at original address'
    Assert-Equal 8 $mark.ApplyVisibleBackgroundColor($target, '#18be6a') 'Marked count'
    Assert-Equal -4142 $sheet.Range('D4').Interior.Pattern 'Unselected format'
    Write-Output 'PASS overlap deduplication, pipeline rounding, formula position, unselected values/formats'

    $sheet.Range('J1').Value2 = 9.876
    $sheet.Range('L1').FormulaArray = '=SUM(H1:H2)'
    $sheet.Range('J1,L1').Select() | Out-Null
    $target = $hostAdapter.CaptureTarget()
    $plan = [AccuX.Core.Operations.RangeWritePlan]::new($target)
    $plan.AddValue(0, 0, 99.0, $null, 0)
    $plan.AddValue(0, 0, 99.0, $null, 1)
    Assert-Equal $false $hostAdapter.ValidateWrite($target, $plan).CanWrite 'Later array area rejected'
    $failed = $false
    try { $hostAdapter.Write($target, $plan) } catch { $failed = $true }
    Assert-Equal $true $failed 'Direct write rechecks all areas'
    Assert-Equal 9.876 $sheet.Range('J1').Value2 'Earlier area not modified'
    Write-Output 'PASS later-area validation failure leaves earlier area unchanged'

    $sheet = $book.Worksheets.Add()
    $sheet.Range('A1').Value2 = 'Group'
    $sheet.Range('B1').Value2 = 'Amount'
    foreach ($i in 2..5) {
        $sheet.Range("A$i").Value2 = $(if ($i % 2 -eq 0) { 'hide' } else { 'show' })
        $sheet.Range("B$i").Value2 = [double]$i
        $sheet.Range("D$i").Value2 = [double]$i
    }
    $sheet.Range('A1:B5').AutoFilter(1, 'show') | Out-Null
    $sheet.Range('B:B').EntireColumn.Hidden = $true
    $sheet.Range('B2:B5,D2:D5').Select() | Out-Null
    $target = $hostAdapter.CaptureTarget()
    $visible = @($hostAdapter.Read($target).Cells | Where-Object { -not $_.IsHidden })
    Assert-Equal 2 $visible.Count 'Filtered rows and hidden column'
    Assert-Equal 2 $mark.ApplyVisibleBackgroundColor($target, '#18be6a') 'Visible-only multi-area marking'
    Assert-Equal -4142 $sheet.Range('B3').Interior.Pattern 'Hidden column unchanged'
    Assert-Equal -4142 $sheet.Range('D2').Interior.Pattern 'Filtered row unchanged'
    Write-Output 'PASS multi-area filtered rows and hidden columns'

    $sheet.Range('D:D,Z100').Select() | Out-Null
    $target = $hostAdapter.CaptureTarget()
    Assert-Equal 5 $target.CellCount 'UsedRange clipping and empty area removal'
    Write-Output 'PASS whole-column clipping and empty-area removal'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $app) {
        $app.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}
