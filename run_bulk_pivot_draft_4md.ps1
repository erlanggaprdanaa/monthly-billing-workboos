. (Join-Path $PSScriptRoot 'billing-config.ps1')

$script = Join-Path $WorkspaceRoot "create_excel_pivot.py"
$blankMergeScript = Join-Path $WorkspaceRoot "merge_blank_to_shared.py"
$sourceWorkbook = $Draft4MD
$outputWorkbook = $Output4MD

Copy-Item $sourceWorkbook $outputWorkbook -Force

& $PyLauncher $PyVersionArgs $blankMergeScript --workbook $outputWorkbook
if ($LASTEXITCODE -ne 0) {
    throw "Blank merge step failed."
}

$mappings = @(
    @{ SourceSheet = "Detail_Astra"; PivotSheet = "Astra" }
    @{ SourceSheet = "Detail_Daya"; PivotSheet = "Daya" }
    @{ SourceSheet = "Detail_MPM"; PivotSheet = "MPM" }
    @{ SourceSheet = "Detail_Wahana"; PivotSheet = "Wahana" }
    @{ SourceSheet = "Detail_Loyalty"; PivotSheet = "Loyalty" }
    @{ SourceSheet = "Detail_Analytic"; PivotSheet = "Analytic" }
    @{ SourceSheet = "Detail_Shared"; PivotSheet = "Shared" }
)

foreach ($mapping in $mappings) {
    Write-Host "Creating pivot '$($mapping.PivotSheet)' from '$($mapping.SourceSheet)'"
    & $PyLauncher $PyVersionArgs $script `
        --source $outputWorkbook `
        --source-sheet $mapping.SourceSheet `
        --output $outputWorkbook `
        --pivot-sheet $mapping.PivotSheet

    if ($LASTEXITCODE -ne 0) {
        throw "Pivot creation failed for source sheet '$($mapping.SourceSheet)'."
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$openedWorkbook = $null

try {
    Write-Host "Removing Split-Costing sheet from workbook"
    $openedWorkbook = $excel.Workbooks.Open($outputWorkbook)
    try {
        $openedWorkbook.Worksheets("Split-Costing").Delete()
    }
    catch {
        Write-Host "Skip delete Split-Costing"
    }
    $openedWorkbook.Save()
}
finally {
    if ($openedWorkbook -ne $null) {
        $openedWorkbook.Close($true)
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($openedWorkbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done. Output workbook: $outputWorkbook"
