from __future__ import annotations

import argparse
import calendar
import json
import math
import re
import subprocess
import tempfile
import shutil
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_FLOOR
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook


AZ_CLI = r"C:\Program Files (x86)\Microsoft SDKs\Azure\CLI2\wbin\az.cmd"
DEFAULT_CACHE_DIR = Path(__file__).resolve().parent / "cache"


OUTPUT_HEADERS = [
    "Period",
    "Agreement Name",
    "Company Name",
    "Tenant Primary Domain",
    "Publisher Subscription ID",
    "Publisher Subscription Name",
    "Product ID",
    "Product Name",
    "Category",
    "Subcategory",
    "Meter Name",
    "Usage Start Date",
    "Usage End Date",
    "Quantity",
    "Unit",
    "Currency",
    "Resource Group Name",
    "Resource Name",
    "Location",
    "Tags_JSON",
    "Tags_KeyValue",
    "Tags_Count",
    "Total",
    "MD Name",
]

VALIDATION_HEADERS = [
    "Source Row",
    "Reservation ID",
    "Reservation Order ID",
    "Product Name",
    "SKU Name",
    "Used Hours",
    "Reserved Hours",
    "Utilization %",
    "Source Total",
    "Allocated Total",
    "Difference",
    "Status",
]

TOTAL_9MD_AND_4MD_HEADERS = [
    "Group",
    "MD Name",
    "Allocated Total",
]

IDR_NUMBER_FORMAT = '"Rp" #,##0'
QUANTITY_NUMBER_FORMAT = '[=0]0.00;[<1]0.00;0.0#'

TOTAL_9MD_MDS = [
    "Anper",
    "CDN",
    "HPM",
    "ITC",
    "MenaraAgung",
    "MSK",
    "Sinsen",
    "TDM",
    "TRIO",
    "Zenobia",
]

TOTAL_4MD_MDS = [
    "Astra",
    "Daya",
    "MPM",
    "Loyalty",
    "Shared",
    "Wahana",
    "Analytic",
]

MONTH_NAME_TO_NUMBER = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}


@dataclass
class ReservationSourceRow:
    row_number: int
    period: str
    period_year: int
    agreement_name: str
    company_name: str
    tenant_primary_domain: str
    publisher_subscription_id: str
    publisher_subscription_name: str
    product_id: Any
    sku_name: str
    product_name: str
    currency: str
    quantity: Decimal
    reseller_sales_unit_price: Decimal
    total_cost: Decimal
    reservation_id: str
    billing_frequency: str
    term_and_billing_cycle: str
    reservation_order_id: str


def infer_location_from_sku_name(sku_name: str) -> str:
    sku_lower = sku_name.lower()
    if "ap southeast" in sku_lower:
        return "southeastasia"
    if "id central" in sku_lower:
        return "indonesiacentral"
    return ""


def load_env_mapping(env_path: Path) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        md_name = key.replace("Detail_", "", 1).strip()
        for rg in value.split(","):
            normalized = rg.strip().lower()
            if normalized:
                mapping[normalized] = md_name
    return mapping


def parse_decimal(value: Any) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value))


def parse_order_id(raw_value: Any) -> str:
    if raw_value is None:
        return ""
    text = str(raw_value).strip()
    if not text:
        return ""
    if text.startswith("{"):
        try:
            payload = json.loads(text)
        except json.JSONDecodeError:
            return text
        return str(payload.get("RIOderId", text))
    return text


def to_source_row(header_map: dict[str, int], row: tuple[Any, ...], row_number: int) -> ReservationSourceRow:
    def get(name: str) -> Any:
        idx = header_map[name]
        return row[idx]

    return ReservationSourceRow(
        row_number=row_number,
        period=str(get("Period")),
        period_year=int(get("Period Year")),
        agreement_name=str(get("Agreement Name")),
        company_name=str(get("Company Name")),
        tenant_primary_domain=str(get("Tenant Primary Domain")),
        publisher_subscription_id=str(get("Publisher Subscription ID")),
        publisher_subscription_name=str(get("Publisher Subscription Name")),
        product_id=get("Product ID"),
        sku_name=str(get("SKU Name")),
        product_name=str(get("Product Name")),
        currency=str(get("Currency")),
        quantity=parse_decimal(get("Quantity")),
        reseller_sales_unit_price=parse_decimal(get("Reseller Sales Unit Price")),
        total_cost=parse_decimal(get("Reseller Sales Price Total")),
        reservation_id=str(get("Reservation ID")),
        billing_frequency=str(get("Billing Frequency")),
        term_and_billing_cycle=str(get("Term and Billing Cycle")),
        reservation_order_id=parse_order_id(get("ReservationOrderId")),
    )


def run_az_json(args: list[str]) -> Any:
    completed = subprocess.run(
        [AZ_CLI, *args, "--only-show-errors", "-o", "json"],
        check=True,
        capture_output=True,
        text=True,
    )
    return json.loads(completed.stdout)


def sanitize_cache_part(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", value)


def load_cache_json(cache_path: Path) -> Any:
    return json.loads(cache_path.read_text(encoding="utf-8"))


def save_cache_json(cache_path: Path, payload: Any) -> None:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    cache_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")


def month_date_range(period: str, year: int) -> tuple[str, str]:
    month_number = MONTH_NAME_TO_NUMBER[period]
    last_day = calendar.monthrange(year, month_number)[1]
    return f"{year:04d}-{month_number:02d}-01", f"{year:04d}-{month_number:02d}-{last_day:02d}"


def extract_rg_and_name(instance_id: str) -> tuple[str, str]:
    parts = instance_id.split("/")
    return parts[4], parts[-1]


def fetch_monthly_summary(
    order_id: str,
    reservation_id: str,
    period: str,
    year: int,
    cache_dir: Path,
    refresh_cache: bool,
) -> dict[str, Any]:
    target_prefix = f"{year:04d}-{MONTH_NAME_TO_NUMBER[period]:02d}-01"
    cache_name = (
        f"summary_{sanitize_cache_part(order_id)}_"
        f"{sanitize_cache_part(reservation_id)}_{year:04d}-{MONTH_NAME_TO_NUMBER[period]:02d}.json"
    )
    cache_path = cache_dir / cache_name
    if cache_path.exists() and not refresh_cache:
        print(f"[CACHE HIT] Summary {reservation_id} {year:04d}-{MONTH_NAME_TO_NUMBER[period]:02d}")
        summaries = load_cache_json(cache_path)
    else:
        print(f"[CACHE MISS] Summary {reservation_id} {year:04d}-{MONTH_NAME_TO_NUMBER[period]:02d}")
        summaries = run_az_json(
            [
                "consumption",
                "reservation",
                "summary",
                "list",
                "--reservation-order-id",
                order_id,
                "--reservation-id",
                reservation_id,
                "--grain",
                "monthly",
            ]
        )
        save_cache_json(cache_path, summaries)
    for summary in summaries:
        if str(summary.get("usageDate", "")).startswith(target_prefix):
            return summary
    raise ValueError(f"Monthly summary not found for reservation {reservation_id} in {period} {year}.")


def fetch_reservation_details(
    order_id: str,
    reservation_id: str,
    start_date: str,
    end_date: str,
    cache_dir: Path,
    refresh_cache: bool,
) -> list[dict[str, Any]]:
    cache_name = (
        f"detail_{sanitize_cache_part(order_id)}_"
        f"{sanitize_cache_part(reservation_id)}_{sanitize_cache_part(start_date)}_{sanitize_cache_part(end_date)}.json"
    )
    cache_path = cache_dir / cache_name
    if cache_path.exists() and not refresh_cache:
        print(f"[CACHE HIT] Detail {reservation_id} {start_date}..{end_date}")
        return load_cache_json(cache_path)

    print(f"[CACHE MISS] Detail {reservation_id} {start_date}..{end_date}")
    details = run_az_json(
        [
            "consumption",
            "reservation",
            "detail",
            "list",
            "--reservation-order-id",
            order_id,
            "--reservation-id",
            reservation_id,
            "--start-date",
            start_date,
            "--end-date",
            end_date,
        ]
    )
    save_cache_json(cache_path, details)
    return details


def allocate_costs(total_cost: Decimal, used_hours_by_resource: list[Decimal]) -> list[Decimal]:
    raw_costs = [hours / sum(used_hours_by_resource) * total_cost for hours in used_hours_by_resource]
    floors = [cost.quantize(Decimal("1"), rounding=ROUND_FLOOR) for cost in raw_costs]
    allocated = floors[:]
    remainder = int(total_cost - sum(floors))
    order = sorted(range(len(raw_costs)), key=lambda idx: (raw_costs[idx] - floors[idx], raw_costs[idx]), reverse=True)
    for idx in order[:remainder]:
        allocated[idx] += Decimal("1")
    return allocated


def normalize_location(value: str) -> str:
    return value.lower().replace(" ", "")


def fetch_resource_location(instance_id: str, cache: dict[str, str], fallback_location: str) -> str:
    # Fast path: for reservation reporting, region from SKU/source reservation
    # is usually sufficient and avoids hundreds of per-resource Azure calls.
    if instance_id not in cache:
        cache[instance_id] = fallback_location
    return cache[instance_id]


def calculate_utilization_percentage(used_hours: Decimal, reserved_hours: Decimal) -> float:
    if reserved_hours == 0:
        return 0.0
    return float((used_hours / reserved_hours).quantize(Decimal("0.0001")))


def build_output_rows(
    source_row: ReservationSourceRow,
    env_mapping: dict[str, str],
    location_cache: dict[str, str],
    progress_label: str,
    cache_dir: Path,
    refresh_cache: bool,
) -> tuple[list[list[Any]], dict[str, Any]]:
    print(
        f"[INFO] {progress_label} Processing row {source_row.row_number} | "
        f"Reservation ID: {source_row.reservation_id} | "
        f"Product: {source_row.product_name}"
    )
    start_date, end_date = month_date_range(source_row.period, source_row.period_year)
    print(f"[INFO]   Fetching monthly summary for {source_row.period} {source_row.period_year}")
    summary = fetch_monthly_summary(
        source_row.reservation_order_id,
        source_row.reservation_id,
        source_row.period,
        source_row.period_year,
        cache_dir,
        refresh_cache,
    )
    print(f"[INFO]   Fetching reservation details from {start_date} to {end_date}")
    details = fetch_reservation_details(
        source_row.reservation_order_id,
        source_row.reservation_id,
        start_date,
        end_date,
        cache_dir,
        refresh_cache,
    )
    used_hours_total = Decimal(str(summary["usedHours"]))
    reserved_hours_total = Decimal(str(summary["reservedHours"]))

    if used_hours_total == 0:
        print(f"[INFO]   No usage found for this reservation in the selected month")
        validation_row = {
            "source_row": source_row.row_number,
            "reservation_id": source_row.reservation_id,
            "reservation_order_id": source_row.reservation_order_id,
            "product_name": source_row.product_name,
            "sku_name": source_row.sku_name,
            "used_hours": float(used_hours_total),
            "reserved_hours": float(reserved_hours_total),
            "utilization_percentage": calculate_utilization_percentage(used_hours_total, reserved_hours_total),
            "source_total": int(source_row.total_cost),
            "allocated_total": 0,
            "difference": int(source_row.total_cost),
            "status": "NO_USAGE",
        }
        return [], validation_row

    if not details:
        raise ValueError(f"No reservation detail rows found for reservation {source_row.reservation_id} despite non-zero usage.")

    hours_per_reserved_quantity = reserved_hours_total / source_row.quantity

    aggregated: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    for detail in details:
        aggregated[detail["instanceId"]] += Decimal(str(detail["usedHours"]))

    sorted_items = sorted(aggregated.items(), key=lambda item: item[1], reverse=True)
    allocated_costs = allocate_costs(source_row.total_cost, [used for _, used in sorted_items])
    print(f"[INFO]   Found {len(sorted_items)} consumer resource(s)")

    output_rows: list[list[Any]] = []
    fallback_location = infer_location_from_sku_name(source_row.sku_name)
    for (instance_id, used_hours), allocated_cost in zip(sorted_items, allocated_costs):
        resource_group, resource_name = extract_rg_and_name(instance_id)
        output_rows.append(
            [
                source_row.period,
                "Microsoft CSP Azure Plan",
                "Divisi Inovasi",
                "divisiinovasi.onmicrosoft.com",
                "a4cdaa2e-caa7-440f-b849-928a3ceb7e31",
                "AzureCSP_DivisiInovasi",
                source_row.reservation_id,
                source_row.product_name,
                f"Reserved Instance {source_row.product_name}",
                source_row.sku_name,
                source_row.sku_name,
                start_date,
                end_date,
                float((used_hours / hours_per_reserved_quantity).quantize(Decimal("0.01"))),
                float(source_row.reseller_sales_unit_price),
                "IDR",
                resource_group,
                resource_name,
                fetch_resource_location(instance_id, location_cache, fallback_location),
                "Reserved Instance",
                source_row.term_and_billing_cycle,
                1,
                int(allocated_cost),
                env_mapping.get(resource_group.lower(), ""),
            ]
        )
    allocated_total = sum(int(row[22]) for row in output_rows)
    print(
        f"[INFO]   Allocated total: {allocated_total:,} | "
        f"Source total: {int(source_row.total_cost):,}"
    )
    validation_row = {
        "source_row": source_row.row_number,
        "reservation_id": source_row.reservation_id,
        "reservation_order_id": source_row.reservation_order_id,
        "product_name": source_row.product_name,
        "sku_name": source_row.sku_name,
        "used_hours": float(used_hours_total),
        "reserved_hours": float(reserved_hours_total),
        "utilization_percentage": calculate_utilization_percentage(used_hours_total, reserved_hours_total),
        "source_total": int(source_row.total_cost),
        "allocated_total": allocated_total,
        "difference": int(source_row.total_cost) - allocated_total,
        "status": (
            "MATCH"
            if int(source_row.total_cost) == allocated_total
            else "PARTIAL_USAGE"
            if allocated_total > 0
            else "MISMATCH"
        ),
    }
    return output_rows, validation_row


def write_output(output_path: Path, rows: list[list[Any]], validation_rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ReservedInstance"
    sheet.append(OUTPUT_HEADERS)
    for row in rows:
        sheet.append(row)

    for row_idx in range(2, sheet.max_row + 1):
        sheet.cell(row_idx, 14).number_format = QUANTITY_NUMBER_FORMAT
        sheet.cell(row_idx, 15).number_format = IDR_NUMBER_FORMAT
        sheet.cell(row_idx, 23).number_format = IDR_NUMBER_FORMAT

    validation_sheet = workbook.create_sheet("Validation")
    validation_sheet.append(VALIDATION_HEADERS)
    for item in validation_rows:
        validation_sheet.append(
            [
                item["source_row"],
                item["reservation_id"],
                item["reservation_order_id"],
                item["product_name"],
                item["sku_name"],
                item["used_hours"],
                item["reserved_hours"],
                item["utilization_percentage"],
                item["source_total"],
                item["allocated_total"],
                item["difference"],
                item["status"],
            ]
        )

    total_source = sum(item["source_total"] for item in validation_rows)
    total_allocated = sum(item["allocated_total"] for item in validation_rows)
    total_diff = total_source - total_allocated
    total_used_hours = sum(item["used_hours"] for item in validation_rows)
    total_reserved_hours = sum(item["reserved_hours"] for item in validation_rows)
    total_utilization = (total_used_hours / total_reserved_hours) if total_reserved_hours else 0
    validation_sheet.append([])
    validation_sheet.append(
        [
            "GRAND TOTAL",
            "",
            "",
            "",
            "",
            total_used_hours,
            total_reserved_hours,
            round(total_utilization, 2),
            total_source,
            total_allocated,
            total_diff,
            "MATCH" if total_diff == 0 else "MISMATCH",
        ]
    )

    for row_idx in range(2, validation_sheet.max_row + 1):
        validation_sheet.cell(row_idx, 6).number_format = "0.00"
        validation_sheet.cell(row_idx, 7).number_format = "0.00"
        validation_sheet.cell(row_idx, 8).number_format = "0.00%"
        validation_sheet.cell(row_idx, 9).number_format = IDR_NUMBER_FORMAT
        validation_sheet.cell(row_idx, 10).number_format = IDR_NUMBER_FORMAT
        validation_sheet.cell(row_idx, 11).number_format = IDR_NUMBER_FORMAT

    totals_sheet = workbook.create_sheet("Total 9-MD&4MD")
    totals_sheet.append(TOTAL_9MD_AND_4MD_HEADERS)

    totals_by_md: dict[str, int] = defaultdict(int)
    for row in rows:
        md_name = "" if row[23] is None else str(row[23]).strip()
        if not md_name:
            continue
        totals_by_md[md_name] += int(row[22])

    def append_group(group_name: str, md_names: list[str]) -> None:
        group_total = 0
        for md_name in md_names:
            md_total = totals_by_md.get(md_name, 0)
            totals_sheet.append([group_name, md_name, md_total])
            group_total += md_total
        totals_sheet.append([group_name, "GRAND TOTAL", group_total])
        totals_sheet.append([])

    append_group("9-MD", TOTAL_9MD_MDS)
    append_group("4-MD", TOTAL_4MD_MDS)

    for row_idx in range(2, totals_sheet.max_row + 1):
        totals_sheet.cell(row_idx, 3).number_format = IDR_NUMBER_FORMAT

    workbook.save(output_path)


def load_source_workbook(source_path: Path):
    try:
        return load_workbook(source_path, data_only=True)
    except PermissionError:
        temp_copy = Path(tempfile.gettempdir()) / f"{source_path.stem}-copy{source_path.suffix}"
        shutil.copy2(source_path, temp_copy)
        return load_workbook(temp_copy, data_only=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate normalized Reserved Instance allocation workbook from Azure historical metrics.")
    parser.add_argument("--source-workbook", required=True, type=Path, help="Reserved instance source workbook.")
    parser.add_argument("--env-file", required=True, type=Path, help=".env mapping file.")
    parser.add_argument("--output-workbook", required=True, type=Path, help="Output workbook path.")
    parser.add_argument("--sheet-name", default="Sheet1", help="Source sheet name.")
    parser.add_argument("--sample-rows", help="Comma-separated Excel row numbers to process for testing.")
    parser.add_argument("--cache-dir", default=DEFAULT_CACHE_DIR, type=Path, help="Directory for Azure API cache files.")
    parser.add_argument("--refresh-cache", action="store_true", help="Ignore existing cache and fetch fresh data from Azure.")
    args = parser.parse_args()

    print(f"[INFO] Starting reserved instance report generation")
    print(f"[INFO] Source workbook : {args.source_workbook}")
    print(f"[INFO] Env mapping     : {args.env_file}")
    print(f"[INFO] Output workbook : {args.output_workbook}")
    print(f"[INFO] Cache directory : {args.cache_dir}")
    if args.refresh_cache:
        print("[INFO] Cache mode      : refresh")
    else:
        print("[INFO] Cache mode      : use cache when available")
    args.cache_dir.mkdir(parents=True, exist_ok=True)
    workbook = load_source_workbook(args.source_workbook)
    sheet = workbook[args.sheet_name]
    rows = list(sheet.iter_rows(values_only=True))
    header_map = {str(name): idx for idx, name in enumerate(rows[0])}

    selected_rows: set[int] | None = None
    if args.sample_rows:
        selected_rows = {int(item.strip()) for item in args.sample_rows.split(",") if item.strip()}

    env_mapping = load_env_mapping(args.env_file)
    location_cache: dict[str, str] = {}
    output_rows: list[list[Any]] = []
    validation_rows: list[dict[str, Any]] = []
    selected_row_numbers: list[int] = []

    for row_number, row in enumerate(rows[1:], start=2):
        if not any(value is not None and value != "" for value in row):
            continue
        if selected_rows is not None and row_number not in selected_rows:
            continue
        selected_row_numbers.append(row_number)

    total_selected = len(selected_row_numbers)
    processed_index = 0

    for row_number, row in enumerate(rows[1:], start=2):
        if row_number not in selected_row_numbers:
            continue
        processed_index += 1
        source_row = to_source_row(header_map, row, row_number)
        if not source_row.reservation_id or not source_row.reservation_order_id:
            print(f"[WARN] Skipping row {row_number}: missing reservation identifiers")
            continue
        progress_label = f"[{processed_index}/{total_selected}]"
        reservation_rows, validation_row = build_output_rows(
            source_row,
            env_mapping,
            location_cache,
            progress_label,
            args.cache_dir,
            args.refresh_cache,
        )
        output_rows.extend(reservation_rows)
        validation_rows.append(validation_row)

    if not output_rows:
        raise ValueError("No output rows were produced. Check sample rows or source workbook data.")

    print(f"[INFO] Writing workbook with {len(output_rows)} output row(s) from {total_selected} source row(s)")
    write_output(args.output_workbook, output_rows, validation_rows)
    print(f"[DONE] Wrote {len(output_rows)} rows to {args.output_workbook}")


if __name__ == "__main__":
    main()
