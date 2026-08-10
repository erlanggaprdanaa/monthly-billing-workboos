$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'billing-config.ps1')

$python = $PythonRuntime
$script = Join-Path $WorkspaceRoot 'append_googlesecops_to_draft.py'
$source = $GoogleSecOpsSource
$draft9 = $Draft9MD
$draft4 = $Draft4MD

& $python $script `
  --source $source `
  --source-sheet '9-MD' `
  --target $draft9 `
  --target-sheets 'Detail_Anper,Detail_Sinsen,Detail_CDN,Detail_MenaraAgung,Detail_MSK,Detail_HPM,Detail_ITC,Detail_TDM,Detail_TRIO'

& $python $script `
  --source $source `
  --source-sheet '4-MD' `
  --target $draft4 `
  --target-sheets 'Detail_Astra,Detail_Daya,Detail_MPM,Detail_Wahana'

& $python $script `
  --source $source `
  --source-sheet 'Shared-9MD' `
  --target $draft9 `
  --target-sheets 'Detail_Zenobia'

Write-Host 'Done.'
