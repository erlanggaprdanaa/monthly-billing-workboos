# Assessment RI Coverage Runbook

Dokumen ini menjelaskan command dan logic perhitungan untuk workbook:

`H:\Work\Code-base\monthly-billing-workbooks\report\assesment\assesment-reservation-planning-2026.xlsx`

## Get Data Forecast RI Expiry

```powershell
py -3 -u "H:\Work\Code-base\monthly-billing-workbooks\report\assesment\forecast_ri_expiry_impact.py" `
  --reservation-source "H:\Work\Code-base\monthly-billing-workbooks\06-June\reserved-instance-june.xlsx" `
  --reserved-report-workbook "H:\Work\Code-base\monthly-billing-workbooks\06-June\reserved-instance-report-june-clean.xlsx" `
  --payg-price-workbook "H:\Work\Code-base\monthly-billing-workbooks\payg-price.xlsx" `
  --draft-workbook "H:\Work\Code-base\monthly-billing-workbooks\06-June\4-MD.xlsx" `
  --draft-workbook "H:\Work\Code-base\monthly-billing-workbooks\06-June\Draft-9MD.xlsx" `
  --expiry-start 2026-08-01 `
  --expiry-end 2026-08-31 `
  --excel-output "H:\Work\Code-base\monthly-billing-workbooks\report\assesment\assesment-reservation-planning-2026.xlsx" `
  --excel-sheet "Forecast RI Expiry"
```

Behavior:

- Sheet `Forecast RI Expiry` akan dibersihkan dulu.
- Data baru ditulis ulang dari hasil kalkulasi terbaru.
- Tidak append, jadi aman untuk rerun.

## Get Data Coverage and Recommendation

```powershell
py -3 -u "H:\Work\Code-base\monthly-billing-workbooks\report\assesment\report_aks_ri_coverage.py" `
  --reserved-instance-report "H:\Work\Code-base\monthly-billing-workbooks\06-June\reserved-instance-report-june-clean.xlsx" `
  --draft-4md "H:\Work\Code-base\monthly-billing-workbooks\06-June\4-MD.xlsx" `
  --draft-9md "H:\Work\Code-base\monthly-billing-workbooks\06-June\Draft-9MD.xlsx" `
  --excel-output "H:\Work\Code-base\monthly-billing-workbooks\report\assesment\assesment-reservation-planning-2026.xlsx"
```

Default output sheet:

| Sheet Name | Isi |
| --- | --- |
| `Coverage` | Coverage RI per AKS |
| `Recommendation in Detail` | Rekomendasi RI per AKS dan VM Size |
| `Summary Reserved` | Summary rekomendasi RI per VM Size |

Behavior:

- Sheet yang ditulis akan dibersihkan dulu.
- Data baru ditulis ulang.
- Tidak append, jadi aman untuk rerun.

## Sheet Selector

Generate only sheet tertentu:

```powershell
--only-sheet "Coverage"
--only-sheet "Recommendation in Detail"
--only-sheet "Summary Reserved"
```

Skip sheet tertentu:

```powershell
--ignore-sheet "Coverage"
```

Contoh generate hanya recommendation detail dan summary:

```powershell
--only-sheet "Recommendation in Detail" `
--only-sheet "Summary Reserved"
```

## Forecast RI Expiry Logic

Output sheet: `Forecast RI Expiry`

| Kolom | Source | Cara Ambil |
| --- | --- | --- |
| `Expiry Date` | `reserved-instance-june.xlsx` | Ambil dari `Subscription End Date` untuk `Reservation ID` yang expired dalam range `--expiry-start` sampai `--expiry-end` |
| `Reservation ID` | `reserved-instance-june.xlsx` + `reserved-instance-report-june-clean.xlsx` | Match `Reservation ID` dari source dengan `Product ID` di report clean |
| `AKS` | `reserved-instance-report-june-clean.xlsx` | Ambil dari `Resource Group Name`, contoh `MC_RG-XXX_CAPSICUM_SOUTHEASTASIA` menjadi `Capsicum` |
| `MD` | `reserved-instance-report-june-clean.xlsx` | Ambil dari kolom `MD Name` |
| `VM Size` | `reserved-instance-june.xlsx` | Parse dari `SKU Name`, contoh `Reserved VM Instance, Standard_D4s_v3, AP Southeast, 1 Year` menjadi `Standard_D4s_v3` |
| `Expiring RI Qty` | `reserved-instance-report-june-clean.xlsx` | Total `Quantity` untuk RI expired tersebut, grouped by `Expiry Date + AKS + VM Size` |
| `PAYG Monthly Price` | `payg-price.xlsx` | Lookup by `VM Size` + `Region` dari sheet `RI PAYG Price` |
| `Current PAYG Left` | `4-MD.xlsx` + `Draft-9MD.xlsx` | Ambil row PAYG dari semua `Detail_*`, filter `Resource Group Name` diawali `MC_`, `Category = Virtual Machines`, exclude `Tags_JSON = Reserved Instance` |
| `Current Coverage` | Calculated | `Current RI Cost / (Current RI Cost + Current PAYG Left)` |
| `After Expired RI Coverage` | Calculated | `(Current RI Cost - Expiring RI Cost) / (Current RI Cost + Current PAYG Left)` |
| `Forecast PAYG Risk` | Calculated | `Current PAYG Left + (Expiring RI Qty × PAYG Monthly Price)` |
| `Used By Resource Name` | `reserved-instance-report-june-clean.xlsx` | Ambil resource name yang menggunakan RI expired tersebut |

Grouping:

`Expiry Date + AKS + VM Size`

Artinya kalau satu AKS punya RI yang expired di dua tanggal berbeda, data akan dipisah menjadi dua row.

## Coverage Logic

Output sheet: `Coverage`

| Kolom | Source | Cara Ambil |
| --- | --- | --- |
| `AKS Name` | `reserved-instance-report-june-clean.xlsx` | Ambil dari `Resource Group Name` yang diawali `MC_` |
| `PAYG Left` | `4-MD.xlsx` + `Draft-9MD.xlsx` | Total PAYG dari `Detail_*`, filter `Resource Group Name` yang match AKS resource group, `Category = Virtual Machines`, exclude Reserved Instance |
| `Coverage %` | Calculated | `RI Total / (RI Total + PAYG Left)` |

Rumus:

```text
Coverage % = RI Total / (RI Total + PAYG Left)
```

Jika `PAYG Left = 0`, maka coverage dianggap `100%`.

## Recommendation in Detail Logic

Output sheet: `Recommendation in Detail`

| Kolom | Source | Cara Ambil |
| --- | --- | --- |
| `AKS Name` | Azure live VMSS + draft workbook | AKS name dari resource group `MC_*` |
| `VM Size PAYG` | Azure live VMSS | Match `Resource Name` PAYG ke VMSS live untuk ambil SKU/VM Size |
| `PAYG Total` | `4-MD.xlsx` + `Draft-9MD.xlsx` | Total PAYG per `AKS + VM Size` |
| `PAYG Monthly Price` | `payg-price.xlsx` | Lookup by `VM Size` + `Region` |
| `Recommended Term` | Command arg | Default `1 Year`, bisa diganti dengan `--recommended-term "3 Years"` |
| `Estimated Workload Qty` | Calculated | `PAYG Total / PAYG Monthly Price` |
| `Nearest Qty` | Calculated | `round_half_up(Estimated Workload Qty)` |
| `Coverage Qty` | Calculated | `CEILING(Estimated Workload Qty)` |
| `Recommended Qty` | Calculated | Default pakai `Nearest Qty` |
| `Qty Decision` | Calculated | Decision text dari `Estimated Workload Qty`, `Nearest Qty`, dan `Coverage Qty` |
| `RI Unit Price Acuan` | `reserved-instance-report-june-clean.xlsx` | Ambil unit price RI existing untuk `VM Size + Term`; fallback manual/public jika tidak ada |
| `Estimated RI Cost` | Calculated | `Recommended Qty × RI Unit Price Acuan` |
| `Potential Saving` | Calculated | `PAYG Total - Estimated RI Cost` |
| `Recommended RI` | Calculated | Jika qty > 0: `{qty} x {VM Size} {Term}`, jika qty 0: `Combine with other cluster` |
| `Note` | Calculated | `Round up`, `Partial, round down`, `Combine with other cluster`, atau fallback price note |

Rumus:

```text
Estimated Workload Qty = PAYG Total / PAYG Monthly Price
Nearest Qty = round_half_up(Estimated Workload Qty)
Coverage Qty = CEILING(Estimated Workload Qty)
Recommended Qty = Nearest Qty
Estimated RI Cost = Recommended Qty × RI Unit Price Acuan
Potential Saving = PAYG Total - Estimated RI Cost
```

Rule note:

| Condition | Note |
| --- | --- |
| `Estimated Workload Qty < 0.5` | `Combine with other cluster` |
| `Recommended Qty = Coverage Qty` dan bukan exact | `Nearest qty, higher coverage` |
| `Recommended Qty < Coverage Qty` | `Nearest qty, partial coverage` |
| `Recommended Qty = Estimated Workload Qty` | `Exact` |

## Summary Reserved Logic

Output sheet: `Summary Reserved`

| Kolom | Source | Cara Ambil |
| --- | --- | --- |
| `VM Size PAYG` | From `Recommendation in Detail` | Group by VM Size |
| `Total PAYG` | Calculated | Sum `PAYG Total` per VM Size |
| `Suggested Purchase Qty` | Calculated | Sum dedicated `Suggested Purchase Qty` per VM Size |
| `Combine Estimated Qty` | Calculated | Sum `Estimated Qty` yang note-nya `Combine with other cluster` |
| `Combine Rounded Qty` | Calculated | `round_half_up(Combine Estimated Qty)` |
| `Total Planning Qty` | Calculated | `Suggested Purchase Qty + Combine Rounded Qty` |
| `Cluster` | Calculated | List cluster dedicated dan combine, contoh `Minerva(5), Wynter(0,42)` |

Rumus:

```text
Total PAYG = SUM(PAYG Total by VM Size)
Suggested Purchase Qty = SUM(dedicated Suggested Purchase Qty)
Combine Estimated Qty = SUM(Estimated Qty where Note contains "Combine with other cluster")
Combine Rounded Qty = round_half_up(Combine Estimated Qty)
Total Planning Qty = Suggested Purchase Qty + Combine Rounded Qty
```

## Source Workbook Notes

| Source | Fungsi |
| --- | --- |
| `reserved-instance-june.xlsx` | Source reservation list, termasuk `Reservation ID`, `SKU Name`, `Quantity`, `Term`, dan `Subscription End Date` |
| `reserved-instance-report-june-clean.xlsx` | RI allocation report, berisi pemakaian RI per resource/AKS/MD |
| `4-MD.xlsx` | PAYG detail untuk 4-MD |
| `Draft-9MD.xlsx` | PAYG detail untuk 9-MD |
| `payg-price.xlsx` | Manual PAYG monthly price untuk forecast expiry risk |
| `assesment-reservation-planning-2026.xlsx` | Output workbook assessment |

## Important Notes

- Untuk forecast expiry, `Product ID` di `reserved-instance-report-june-clean.xlsx` harus berisi `Reservation ID`.
- Untuk coverage/recommendation, script membaca Azure live VMSS untuk mapping `Resource Name` ke `VM Size`.
- Pastikan workbook output Excel ditutup sebelum running script, karena Excel yang terbuka akan lock file dan menyebabkan `PermissionError`.
