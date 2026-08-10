# Monthly Billing Workbooks

Workspace kerja saat ini:

`H:\Work\Code-base\monthly-billing-workbooks`

## Config

Semua variable PowerShell utama ada di:

`billing-config.ps1`

Contoh:
- `$MonthFolder`
- `$Draft4MD`
- `$Draft9MD`
- `$Output9MD`
- `$ReservedInstanceSource`
- `$ReservedInstanceOutput`
- `$SummaryTotalWorkbook`
- `$SummarySheetName`
- `$SummaryMonthDate`
- `$PaygTotalExpected`

Kalau ganti bulan, cukup update file `billing-config.ps1`.

## Commands

Load config dulu:

```powershell
. "H:\Work\Code-base\monthly-billing-workbooks\billing-config.ps1"
```

Tambahkan biaya Google SecOps:

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_googlesecops_mapping.ps1"
```

Hitung bagi rata reserved instance:

```powershell
& $PythonRuntime "H:\Work\Code-base\monthly-billing-workbooks\generate_reserved_instance_report.py" --source-workbook $ReservedInstanceSource --env-file $ReservedInstanceEnv --output-workbook $ReservedInstanceOutput
```

Hapus kolom tambahan Draft 9-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\cleanse_draft_workbook.py" --source $Draft9MD
```

Hapus kolom tambahan Draft 4-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\cleanse_draft_workbook.py" --source $Draft4MD
```

Bagi rata Loyalty ke Shared 9-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\rebalance_loyalty_to_zenobia.py" --draft-4md $Draft4MD --draft-9md $Draft9MD --target-total 75000000
```

Bulk pivot 4-MD:

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_bulk_pivot_draft_4md.ps1"
```

Bulk pivot 9-MD:

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_bulk_pivot_draft_9md.ps1"
```

Create Summary Sheet 4-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\create_draft_4md_summary.py" --source $Draft4MD
```

Update `summary-total.xlsx`:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\update_summary_total.py" `
  --summary-workbook $SummaryTotalWorkbook `
  --draft-4md $Output4MD `
  --draft-9md $Draft9MD `
  --reserved-instance-report-workbook $ReservedInstanceOutput `
  --shared-license-workbook $SharedLicenseTemplate `
  --sheet-name $SummarySheetName `
  --month-date $SummaryMonthDate `
  --payg-total-expected $PaygTotalExpected
```
