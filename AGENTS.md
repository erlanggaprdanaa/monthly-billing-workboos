# AGENTS.md

## Purpose

This repository prepares monthly Azure billing workbooks for the 4-MD and 9-MD
customer groups. It also creates Reserved Instance (RI) allocations, summary
workbooks, PAYG checks, and AKS reservation assessments.

This file is the operating guide for contributors and coding agents. It explains
which script owns each step and which data must not be changed accidentally.

## Existing Documentation

Use these documents before changing logic:

| Document | Coverage |
| --- | --- |
| `MONTHLY_BILLING_RUNBOOK.md` | Normal monthly billing process and command order. |
| `WORKFLOW.md` | Short command reference and PowerShell configuration overview. |
| `report/assesment/ASSESSMENT-RI-COVERAGE-RUNBOOK.md` | AKS RI coverage, purchase recommendation, cluster quantity, and expiry forecast logic. |

`AGENTS.md` complements these runbooks. Do not duplicate full command examples
here unless a script is not otherwise documented.

## Working Rules

1. Load `billing-config.ps1` before normal monthly commands. Update its year,
   month, folder, expected PAYG amount, and input/output paths for a new month.
2. `.env` and `.env-nawa` map resource groups to `MD Name`. They contain billing
   mappings and must not be committed or exposed.
3. Close Excel workbooks before scripts write to them. A locked workbook can cause
   `PermissionError` or a partial save.
4. Preserve raw source workbooks. Generated files may be recreated; raw Crayon
   export files must remain unchanged.
5. Currency uses numeric IDR values. Do not convert money to formatted strings.
6. Use `Prorated Reseller Sales Price Total` as the RI source billing amount. It
   includes cancellation credits as negative values.
7. `ReservedInstance` is billable/mappable output. `Validation`, `Cancelled RI
   Usage`, and `RI Reconciliation` are control sheets and must not be mapped into
   a customer draft workbook.
8. A row with `Product ID = TEMP-RI-BALANCE-2026-09` is a September-only manual
   reconciliation adjustment. It is not Azure usage or an Azure Reservation ID.
   Do not copy it into a new month or add it to generator logic.

## Monthly Data Flow

```text
Crayon raw export
  -> Draft-4MD.xlsx / Draft-9MD.xlsx
  -> Google SecOps and RI mapping
  -> 4-MD.xlsx / 9-MD.xlsx pivots
  -> summary-total.xlsx

Crayon RI export + Azure Reservation Consumption API
  -> reserved-instance-report-<month>-clean.xlsx
  -> map RI to Draft-4MD.xlsx and Draft-9MD.xlsx
  -> summary-total.xlsx and AKS assessment workbooks
```

Normal monthly command order is in `MONTHLY_BILLING_RUNBOOK.md`.

## Configuration and PowerShell Entrypoints

| File | Purpose |
| --- | --- |
| `billing-config.ps1` | Main month configuration. Defines paths, Python runtime, input workbooks, output workbooks, expected PAYG, and summary sheet name. |
| `billing-config-nawa.ps1` | Alternative configuration for the Nawa subscription and `.env-nawa` mapping. |
| `run_googlesecops_mapping.ps1` | Calls `append_googlesecops_to_draft.py` for 4-MD, 9-MD, and Shared-9MD source tabs. |
| `run_bulk_pivot_draft_4md.ps1` | Copies Draft-4MD to 4-MD, merges `Blank` into Shared, creates pivots, then removes helper sheets through Excel COM. Windows and Excel are required. |
| `run_bulk_pivot_draft_9md.ps1` | Copies Draft-9MD to 9-MD, creates/reorders pivots, then removes detail and helper sheets through Excel COM. Windows and Excel are required. |

## Core Billing Scripts

| Script | Input | Output / Logic |
| --- | --- | --- |
| `cleanse_draft_workbook.py` | Draft workbook | Removes extra pricing columns from detail sheets and formats `Total` as IDR. |
| `append_googlesecops_to_draft.py` | Google SecOps source + selected source sheet + draft detail sheets | Validates common headers, then appends Google SecOps billing rows to destination detail sheets. |
| `append_shared_license_to_reserved_shared.py` | Shared license template + Draft-4MD | Appends shared license rows to the `Reserved Instance Shared` sheet. |
| `rebalance_loyalty_to_zenobia.py` | Draft-4MD + Draft-9MD + target total | Moves Loyalty rows into 9-MD Zenobia until the requested target is reached. This is a business allocation override, not raw billing transformation. |
| `merge_blank_to_shared.py` | Draft workbook | Appends `Blank` rows into `Detail_Shared`, then removes `Blank`. Used by the 4-MD pivot wrapper. |
| `create_excel_pivot.py` | Workbook + source sheet + pivot sheet name | Creates an Excel PivotTable. It is called by the pivot PowerShell wrappers. |
| `create_draft_4md_summary.py` | Draft-4MD | Creates/refreshes the Summary sheet inside Draft-4MD. |
| `update_summary_total.py` | `summary-total.xlsx`, 4-MD, 9-MD, RI report, shared license template | Refreshes the monthly summary and validates expected vs calculated PAYG and RI totals. |

## Reserved Instance Scripts

| Script | Input | Output / Logic |
| --- | --- | --- |
| `generate_reserved_instance_report.py` | Crayon RI source, `.env` mapping, Azure CLI reservation consumption API | Creates `ReservedInstance`, `Validation`, `Total 9-MD&4MD`, `Cancelled RI Usage`, and `RI Reconciliation`. `ReservedInstance` maps successful RI allocation and positive cancelled usage. Negative cancelled allocation stays out of MD distribution. |
| `map_reserved_instance_to_draft_4md.py` | RI report + Draft-4MD | Deletes earlier RI rows/sheets for the affected target, then maps `ReservedInstance` rows by `MD Name` into 4-MD detail and RI sheets. |
| `map_reserved_instance_to_draft_9md.py` | RI report + Draft-9MD | Equivalent mapping for 9-MD target sheets. |
| `fix_draft_md_sheet_mapping.py` | Draft-4MD + Draft-9MD + `.env` mapping | Moves misplaced MD sheets to the correct draft workbook and removes the stale source sheet. |

### RI Report Rules

- `Product ID` in `ReservedInstance` is the Reservation ID for real Azure RI rows.
- Azure calls use both Reservation Order ID and Reservation ID.
- `--refresh-cache` ignores local `cache/` data and calls Azure again.
- `Cancelled RI Usage` is informational historical Azure utilization; it is not a
  direct source for RI billing totals.
- `RI Reconciliation` must be checked after every rerun. Its source net should
  reconcile to the filtered Crayon RI source amount, subject to documented
  row-level rounding.
- Temporary September balancing rows are intentionally not generated by the
  Python generator and must not be treated as real reservations.

## PAYG and Comparison Scripts

| Script | Input | Output / Logic |
| --- | --- | --- |
| `monthly_check_payg.py` | `.env` resource groups and Nava FinOps API | Retrieves current-month PAYG/cost detail by RG. Uses current execution month and writes/replaces the month sheet in `monthly-check-payg.xlsx`. |
| `compare_monthly_detail.py` | Previous and current workbook + one `Detail_*` sheet | Produces a detailed meter/resource variance report for one MD sheet. |
| `compare_all_details.py` | Previous and current draft workbooks | Runs variance comparison across all detail sheets and creates a combined Markdown report. |
| `compare_summary_total.py` | `summary-total.xlsx` | Compares the selected month with the previous month and prints/creates component-level movements for PAYG, RI, SecOps, and Shared License. |

## Templates and Supporting Builders

| Script | Purpose |
| --- | --- |
| `build_template_summary.py` | Builds the base summary workbook template. Run only when creating or intentionally replacing a template. |
| `build_template_shared_license.py` | Builds the shared-license template used by the monthly append step. |

## AKS Assessment Scripts

All assessment output writes to `report/assesment/assesment-reservation-planning-2026.xlsx`.
Each target sheet is cleared before fresh data is written, so reruns replace old
assessment data rather than append duplicates.

| Script | Output Sheet / Purpose |
| --- | --- |
| `report/assesment/forecast_ri_expiry_impact.py` | `Forecast RI Expiry`. Calculates cost and coverage impact of RI expiry using RI source, clean RI report, PAYG prices, and draft PAYG rows. |
| `report/assesment/report_aks_ri_coverage.py` | `Coverage`, `Recommendation in Detail`, and `Summary Reserved`. Supports `--only-sheet` and `--ignore-sheet`. |
| `report/assesment/cluster_qty_summary.py` | `Cluster QTY Summary`. Derives PAYG VM quantity from PAYG cost divided by PAYG monthly price, adds RI quantity, and includes rounded total quantity. |
| `report_aks_ri_coverage.py` | Earlier root-level coverage report. Prefer the version in `report/assesment/` for the current multi-sheet assessment workflow. |

## Script Status

| Area | Existing Documentation Before This File | Notes |
| --- | --- | --- |
| Normal monthly flow | Yes | `MONTHLY_BILLING_RUNBOOK.md` is authoritative for order and commands. |
| RI generation and distribution | Yes, partial | Existing runbook covered commands; this file adds API, cancellation, reconciliation, and temporary-row boundaries. |
| AKS assessment | Yes | `report/assesment/ASSESSMENT-RI-COVERAGE-RUNBOOK.md` contains formulas and source definitions. |
| Comparison utilities | No consolidated guide | Documented here. |
| PAYG live check | No consolidated guide | Documented here. |
| Template and maintenance utilities | No consolidated guide | Documented here. |

## Safe Change Checklist

Before changing a script:

1. Identify its caller in the PowerShell wrappers and runbooks.
2. Confirm the source columns and target sheet names are unchanged.
3. Do not remove `Status`, `Validation`, or reconciliation behavior from RI logic.
4. Test on a copied/sample workbook first when changing allocations or totals.
5. Reconcile `Prorated Reseller Sales Price Total`, `ReservedInstance`, and
   `RI Reconciliation` after RI changes.
6. Do not include `TEMP-RI-BALANCE-*` rows in next-month data or permanent logic.

