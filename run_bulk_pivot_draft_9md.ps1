. (Join-Path $PSScriptRoot 'billing-config.ps1')

$script = Join-Path $WorkspaceRoot "create_excel_pivot.py"
$sourceWorkbook = $Draft9MD
$outputWorkbook = $Output9MD

$mappings = @(
    @{ SourceSheet = "Detail_CDN"; PivotSheet = "Capella Dinamik Nusantara (CDN)" }
    @{ SourceSheet = "Detail_HPM"; PivotSheet = "HPM" }
    @{ SourceSheet = "Detail_MenaraAgung"; PivotSheet = "Menara Agung" }
    @{ SourceSheet = "Detail_ITC"; PivotSheet = "ITC" }
    @{ SourceSheet = "Detail_TRIO"; PivotSheet = "Trio" }
    @{ SourceSheet = "Detail_Sinsen"; PivotSheet = "SinSen" }
    @{ SourceSheet = "Detail_Anper"; PivotSheet = "AnPer" }
    @{ SourceSheet = "Detail_MSK"; PivotSheet = "MSK" }
    @{ SourceSheet = "Detail_TDM"; PivotSheet = "TDM" }
    @{ SourceSheet = "Detail_Zenobia"; PivotSheet = "Shared" }
)

Copy-Item $sourceWorkbook $outputWorkbook -Force

$firstRun = $true
foreach ($mapping in $mappings) {
    Write-Host "Creating pivot '$($mapping.PivotSheet)' from '$($mapping.SourceSheet)'"
    $sourceForRun = if ($firstRun) { $sourceWorkbook } else { $outputWorkbook }
    & $PyLauncher $PyVersionArgs $script `
        --source $sourceForRun `
        --source-sheet $mapping.SourceSheet `
        --output $outputWorkbook `
        --pivot-sheet $mapping.PivotSheet
    if ($LASTEXITCODE -ne 0) {
        throw "Pivot creation failed for sheet '$($mapping.SourceSheet)'."
    }
    $firstRun = $false
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    Write-Host "Reordering pivot sheets"
    $workbook = $excel.Workbooks.Open($outputWorkbook)
    $anchorSheet = $workbook.Worksheets("Detail_Anper")
    $desiredOrder = @(
        "Capella Dinamik Nusantara (CDN)"
        "HPM"
        "Menara Agung"
        "ITC"
        "Trio"
        "SinSen"
        "AnPer"
        "MSK"
        "TDM"
        "Shared"
    )

    foreach ($pivotName in [System.Linq.Enumerable]::Reverse([string[]]$desiredOrder)) {
        $pivotSheet = $workbook.Worksheets($pivotName)
        $pivotSheet.Move($anchorSheet)
        $anchorSheet = $pivotSheet
    }

    Write-Host "Removing detail sheets from output workbook"
    $detailSheets = @(
        "Detail_Anper"
        "Detail_Sinsen"
        "Detail_CDN"
        "Detail_MenaraAgung"
        "Detail_MSK"
        "Detail_HPM"
        "Detail_ITC"
        "Detail_TDM"
        "Detail_TRIO"
        "Detail_Zenobia"
    )

    foreach ($detailSheetName in $detailSheets) {
        try {
            $workbook.Worksheets($detailSheetName).Delete()
        }
        catch {
            Write-Host "Skip delete $detailSheetName"
        }
    }

    try {
        Write-Host "Removing Split-Costing sheet from output workbook"
        $workbook.Worksheets("Split-Costing").Delete()
    }
    catch {
        Write-Host "Skip delete Split-Costing"
    }

    $workbook.Save()
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($true)
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "Done. Output workbook: $outputWorkbook"
