from __future__ import annotations

import argparse
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

import win32com.client as win32


SOURCE_SHEET = "Detail_Loyalty"
TARGET_SHEET = "Detail_Zenobia"
TOTAL_HEADER = "Total"
XL_UP = -4162
XL_TO_LEFT = -4159


@dataclass
class RowCandidate:
    row_number: int
    total: int
    values: list[Any]


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


def last_used_row(sheet) -> int:
    return sheet.Cells(sheet.Rows.Count, 1).End(XL_UP).Row


def read_sheet_rows(sheet, header_map: dict[str, int]) -> list[RowCandidate]:
    total_col = header_map[TOTAL_HEADER]
    end_row = last_used_row(sheet)
    end_col = max(header_map.values())
    candidates: list[RowCandidate] = []

    for row_idx in range(2, end_row + 1):
        row_values = list(sheet.Range(sheet.Cells(row_idx, 1), sheet.Cells(row_idx, end_col)).Value[0])
        if not any(value is not None and value != "" for value in row_values):
            continue
        total_value = row_values[total_col - 1]
        if total_value in (None, ""):
            continue
        total_int = int(Decimal(str(total_value)))
        if total_int <= 0:
            continue
        candidates.append(RowCandidate(row_number=row_idx, total=total_int, values=row_values))
    return candidates


def choose_rows(candidates: list[RowCandidate], gap: int) -> list[RowCandidate]:
    # Greedy descending: close to target, never exceed gap.
    chosen: list[RowCandidate] = []
    running_total = 0
    for candidate in sorted(candidates, key=lambda item: (item.total, item.row_number), reverse=True):
        if running_total + candidate.total > gap:
            continue
        chosen.append(candidate)
        running_total += candidate.total
        if running_total == gap:
            break
    return sorted(chosen, key=lambda item: item.row_number)


def append_rows(sheet, start_row: int, rows: list[RowCandidate], width: int) -> None:
    if not rows:
        return
    matrix = tuple(tuple(row.values[:width]) for row in rows)
    end_row = start_row + len(rows) - 1
    sheet.Range(sheet.Cells(start_row, 1), sheet.Cells(end_row, width)).Value = matrix


def delete_rows(sheet, row_numbers: list[int]) -> None:
    for row_number in sorted(row_numbers, reverse=True):
        sheet.Rows(row_number).Delete()


def process(draft_4md: Path, draft_9md: Path, target_total: int) -> None:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb4 = None
    wb9 = None
    try:
        print(f"[INFO] Opening source workbook : {draft_4md}")
        wb4 = excel.Workbooks.Open(str(draft_4md))
        print(f"[INFO] Opening target workbook : {draft_9md}")
        wb9 = excel.Workbooks.Open(str(draft_9md))

        source_ws = wb4.Worksheets(SOURCE_SHEET)
        target_ws = wb9.Worksheets(TARGET_SHEET)

        source_headers = get_header_map(source_ws)
        target_headers = get_header_map(target_ws)
        if source_headers != target_headers:
            raise ValueError("Source and target headers do not match.")

        source_rows = read_sheet_rows(source_ws, source_headers)
        target_rows = read_sheet_rows(target_ws, target_headers)
        current_total = sum(row.total for row in target_rows)
        gap = target_total - current_total

        print(f"[INFO] Current {TARGET_SHEET} total : {current_total:,}")
        print(f"[INFO] Target total            : {target_total:,}")
        print(f"[INFO] Gap                     : {gap:,}")

        if gap <= 0:
            print("[INFO] No top-up needed. Target sheet already meets or exceeds target total.")
            return

        selected_rows = choose_rows(source_rows, gap)
        selected_total = sum(row.total for row in selected_rows)
        if not selected_rows:
            raise ValueError("No eligible rows found in Detail_Loyalty for this gap.")

        print(f"[INFO] Selected {len(selected_rows)} row(s) from {SOURCE_SHEET}")
        print(f"[INFO] Selected total : {selected_total:,}")
        print(f"[INFO] Remaining gap  : {gap - selected_total:,}")

        target_start_row = last_used_row(target_ws) + 1
        append_rows(target_ws, target_start_row, selected_rows, max(target_headers.values()))
        delete_rows(source_ws, [row.row_number for row in selected_rows])

        print(f"[INFO] Appended into {TARGET_SHEET} starting at row {target_start_row}")
        print(f"[INFO] Deleted moved rows from {SOURCE_SHEET}")

        wb4.Save()
        wb9.Save()
        print("[DONE] Rebalance completed.")
    finally:
        if wb4 is not None:
            wb4.Close(SaveChanges=True)
        if wb9 is not None:
            wb9.Close(SaveChanges=True)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Move rows from Draft-4MD Detail_Loyalty into Draft-9MD Detail_Zenobia until target total is reached."
    )
    parser.add_argument("--draft-4md", required=True, type=Path, help="Path to Draft-4MD.xlsx")
    parser.add_argument("--draft-9md", required=True, type=Path, help="Path to Draft-9MD.xlsx")
    parser.add_argument("--target-total", type=int, default=75_000_000, help="Target total for Detail_Zenobia")
    args = parser.parse_args()

    process(args.draft_4md, args.draft_9md, args.target_total)


if __name__ == "__main__":
    main()
