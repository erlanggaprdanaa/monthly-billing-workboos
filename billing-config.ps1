$WorkspaceRoot = "H:\Work\Code-base\monthly-billing-workbooks"
$BaseDataRoot = "H:\Work\Code-base\monthly-billing-workbooks"

$BillingYear = "2026"
$MonthNumber = "06"
$MonthName = "June"
$MonthKey = $MonthName.ToLower()
$MonthFolderName = "$MonthNumber-$MonthName"
$MonthFolder = Join-Path $BaseDataRoot $MonthFolderName

$PythonRuntime = "C:\Users\DYS\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$PyLauncher = "py"
$PyVersionArgs = @("-3")

$GoogleSecOpsSource = Join-Path $BaseDataRoot "sheet-googlesecops.xlsx"
$ReservedInstanceEnv = Join-Path $WorkspaceRoot ".env"
$ReservedInstanceSource = Join-Path $MonthFolder "reserved-instance-$MonthKey.xlsx"
$ReservedInstanceOutput = Join-Path $MonthFolder "reserved-instance-report-$MonthKey-clean.xlsx"
$SharedLicenseTemplate = Join-Path $WorkspaceRoot "template-shared-license.xlsx"
$SummaryTotalWorkbook = Join-Path $WorkspaceRoot "summary-total.xlsx"
$PaygTotalExpected = 1006541574
$SummarySheetName = "$MonthName $BillingYear"
$SummaryMonthDate = "$BillingYear-$MonthNumber-01"

$Draft4MD = Join-Path $MonthFolder "Draft-4MD.xlsx"
$Output4MD = Join-Path $MonthFolder "4-MD.xlsx"
$Draft9MD = Join-Path $MonthFolder "Draft-9MD.xlsx"
$Output9MD = Join-Path $MonthFolder "9-MD.xlsx"
