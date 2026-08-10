from __future__ import annotations

import argparse
from pathlib import Path

import win32com.client as win32


COLUMNS_TO_REMOVE = [
    "Unit Price",
    "Price Adjustment",
    "Sales Unit Price",
    "Percentage Markup",
]

ACCOUNTING_IDR_FORMAT = '_-[$Rp-421]* #,##0_-;\\-[$Rp-421]* #,##0_-;_-[$Rp-421]* "-"_-;_-@_-'
ACCOUNTING_IDR_FORMAT_LOCAL = '_-[$Rp-421] * #.##0_-;\\-[$Rp-421] * #.##0_-;_-[$Rp-421] * "-"_-;_-@_-'
XL_TO_LEFT = -4159


def normalize_header(value: object) -> str:
    return "" if value is None else str(value).strip()


def get_header_map(sheet) -> dict[str, int]:
    last_col = sheet.Cells(1, sheet.Columns.Count).End(XL_TO_LEFT).Column
    values = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col)).Value
    header_map: dict[str, int] = {}
    for col_idx in range(1, last_col + 1):
        value = values[0][col_idx - 1] if isinstance(values, tuple) else values
        header = normalize_header(value)
        if header:
            header_map[header] = col_idx
    return header_map


def is_detail_sheet(sheet) -> bool:
    header_map = get_header_map(sheet)
    return "Total" in header_map and "Resource Name" in header_map and "Category" in header_map


def cleanse_sheet(sheet) -> list[str]:
    header_map = get_header_map(sheet)
    removed_headers = [header for header in COLUMNS_TO_REMOVE if header in header_map]

    for header in sorted(removed_headers, key=lambda item: header_map[item], reverse=True):
        sheet.Columns(header_map[header]).Delete()

    header_map = get_header_map(sheet)
    if "Total" not in header_map:
        raise ValueError(f"Sheet '{sheet.Name}' does not have 'Total' column after cleanup.")

    total_col = header_map["Total"]
    total_range = sheet.Columns(total_col)
    total_range.NumberFormat = ACCOUNTING_IDR_FORMAT
    try:
        total_range.NumberFormatLocal = ACCOUNTING_IDR_FORMAT_LOCAL
    except Exception:
        pass

    return removed_headers


def process_workbook(source_path: Path) -> list[str]:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    results: list[str] = []
    try:
        print("[INFO] Opening workbook in Excel")
        workbook = excel.Workbooks.Open(str(source_path))

        for sheet in workbook.Worksheets:
            if not is_detail_sheet(sheet):
                continue
            print(f"[INFO] Cleansing sheet '{sheet.Name}'")
            removed_headers = cleanse_sheet(sheet)
            results.append(
                f"[DONE] {sheet.Name}: removed {removed_headers if removed_headers else 'no optional columns found'}, formatted 'Total'"
            )

        print("[INFO] Saving workbook")
        workbook.Save()
        return results
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Remove extra pricing columns from draft detail sheets and format Total as IDR accounting."
    )
    parser.add_argument("--source", required=True, type=Path, help="Draft workbook path.")
    args = parser.parse_args()

    print(f"[INFO] Source workbook : {args.source}")
    results = process_workbook(args.source)
    for line in results:
        print(line)
    print("[DONE] Workbook cleansing completed.")


if __name__ == "__main__":
    main()
