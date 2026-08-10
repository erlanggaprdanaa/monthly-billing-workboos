# Monthly Billing Runbook

Workspace:

`H:\Work\Code-base\monthly-billing-workbooks`

## Load Config

Jalankan dulu:

```powershell
. "H:\Work\Code-base\monthly-billing-workbooks\billing-config.ps1"
```

## Alur Kerja

### 1. Tambahkan biaya Google SecOps

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_googlesecops_mapping.ps1"
```

### 2. Hapus kolom tambahan di draft workbook

Kolom yang dihapus:
- `Unit Price`
- `Price Adjustment`
- `Sales Unit Price`
- `Percentage Markup`

Draft 9-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\cleanse_draft_workbook.py" --source $Draft9MD
```

Draft 4-MD:

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\cleanse_draft_workbook.py" --source $Draft4MD
```

### 3. Generate report Reserved Instance

Pakai command ini untuk run normal:

```powershell
& $PythonRuntime "H:\Work\Code-base\monthly-billing-workbooks\generate_reserved_instance_report.py" --source-workbook $ReservedInstanceSource --env-file $ReservedInstanceEnv --output-workbook $ReservedInstanceOutput
```

Catatan:
- jika sudah pernah execute 1x untuk bulan yang sama, normalnya tidak perlu refresh
- script akan memakai cache kalau datanya sudah ada

Kalau mau paksa refresh dari Azure dan tidak memakai cache:

```powershell
& $PythonRuntime "H:\Work\Code-base\monthly-billing-workbooks\generate_reserved_instance_report.py" --source-workbook $ReservedInstanceSource --env-file $ReservedInstanceEnv --output-workbook $ReservedInstanceOutput --refresh-cache
```

### 4. Sebar Reserved Instance ke Draft 9-MD

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\map_reserved_instance_to_draft_9md.py" --source-workbook $ReservedInstanceOutput --target-workbook $Draft9MD
```

### 5. Sebar Reserved Instance ke Draft 4-MD

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\map_reserved_instance_to_draft_4md.py" --source-workbook $ReservedInstanceOutput --target-workbook $Draft4MD
```

### 6. Bagi rata Loyalty ke Shared 9-MD

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\rebalance_loyalty_to_zenobia.py" --draft-4md $Draft4MD --draft-9md $Draft9MD --target-total 75000000
```

### 7. Tambahkan Shared License ke Reserved Instance Shared

```powershell
& $PyLauncher $PyVersionArgs "H:\Work\Code-base\monthly-billing-workbooks\append_shared_license_to_reserved_shared.py" --template-workbook $SharedLicenseTemplate --target-workbook $Draft4MD
```

### 8. Generate pivot workbook

Pivot 9-MD:

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_bulk_pivot_draft_9md.ps1"
```

Pivot 4-MD:

Catatan:
- script ini juga append sheet `Blank` ke `Detail_Shared`

```powershell
& "H:\Work\Code-base\monthly-billing-workbooks\run_bulk_pivot_draft_4md.ps1"
```

### 9. Update summary total dan validasi PAYG

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

Hasil yang akan muncul di `summary-total.xlsx`:
- PAYG
- Reserved Instance
- Google SecOps
- Shared Licensi
- PAYG Expected / Calculated / Difference / Status
- RI Expected / Calculated / Difference / Status
- Last Updated
- Update Mode

## Rule Final Summary

### PAYG

#### Draft-4MD

Untuk sheet berikut:
- `Detail_Astra`
- `Detail_Daya`
- `Detail_MPM`
- `Detail_Wahana`

Rumus:
- `PAYG = Overall - SecOps`

Untuk sheet berikut:
- `Detail_Analytic`
- `Detail_Loyalty`

Rumus:
- `PAYG = Overall - RI`

Untuk sheet berikut:
- `Detail_Shared`

Rumus:
- tetap memakai rule saat ini di script

#### Draft-9MD

Untuk semua sheet `Detail_*` di 9-MD:
- `Detail_Anper`
- `Detail_Sinsen`
- `Detail_CDN`
- `Detail_MenaraAgung`
- `Detail_MSK`
- `Detail_HPM`
- `Detail_ITC`
- `Detail_TDM`
- `Detail_TRIO`
- `Detail_Zenobia`

Rumus:
- `PAYG = Overall - RI - SecOps`

### Reserved Instance

#### Draft-4MD

Untuk kolom berikut:
- `Astra`
- `Daya`
- `MPM`
- `Wahana`
- `Shared`

Ambil dari sheet depan:
- `Reserved Instance Astra`
- `Reserved Instance Daya`
- `Reserved Instance MPM`
- `Reserved Instance Wahana`
- `Reserved Instance Shared`

Catatan khusus:
- `Reserved Instance Shared` tidak membawa baris Shared License

Untuk:
- `Analytic`
- `Loyalty`

Ambil dari filter RI di:
- `Detail_Analytic`
- `Detail_Loyalty`

#### Draft-9MD

Ambil dari filter RI di masing-masing `Detail_*`.

### Google SecOps

Google SecOps tetap dihitung terpisah per kolom.

## Formula Excel Summary

### 4-MD PAYG

Formula dinamis:

```excel
=SUM(INDIRECT("'Detail_"&C$3&"'!W2:W50000"))
```

### 4-MD Reserved Instance

Formula dinamis:

```excel
=SUM(INDIRECT("'Reserved Instance "&C$3&"'!W2:W50000"))
```

### Shared PAYG

Gabungan `Detail_Shared` dan `Detail_Loyalty`:

```excel
=SUM(Detail_Shared!W2:W50000)+SUM(Detail_Loyalty!W2:W50000)
```

### Shared Reserved Instance

Nama sheet ada spasi, jadi wajib pakai petik satu:

```excel
=SUM('Reserved Instance Shared'!W2:W50000)
```

### Analytic PAYG

```excel
=SUM(Detail_Analytic!W2:W50000)
```

## Catatan Penting

- Jika nama sheet ada spasi di Excel formula, wajib pakai tanda petik satu
- `summary-total.xlsx` akan di-refresh kalau sheet bulan yang sama dijalankan ulang
- `Last Updated` dipakai sebagai penanda terakhir script dijalankan
- `Update Mode` menunjukkan `Created new sheet` atau `Refreshed existing sheet`
