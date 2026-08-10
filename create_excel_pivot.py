from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path

import win32com.client as win32


XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_SUM = -4157
XL_COMPACT_ROW = 0
ACCOUNTING_IDR_FORMAT = '_-[$Rp-421]* #,##0_-;\\-[$Rp-421]* #,##0_-;_-[$Rp-421]* "-"_-;_-@_-'
ACCOUNTING_IDR_FORMAT_LOCAL = '_-[$Rp-421] * #.##0_-;\\-[$Rp-421] * #.##0_-;_-[$Rp-421] * "-"_-;_-@_-'


def build_pivot(
    source_path: Path,
    source_sheet: str,
    output_path: Path,
    pivot_sheet: str,
    pivot_name: str,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if source_path.resolve() != output_path.resolve():
        shutil.copy2(source_path, output_path)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    try:
        workbook = excel.Workbooks.Open(str(output_path))
        source_ws = workbook.Worksheets(source_sheet)

        try:
            existing_pivot_ws = workbook.Worksheets(pivot_sheet)
            existing_pivot_ws.Delete()
        except Exception:
            pass

        pivot_ws = workbook.Worksheets.Add()
        pivot_ws.Name = pivot_sheet
        pivot_ws.Move(Before=source_ws)
        pivot_ws.Range("A1").Value = pivot_sheet
        pivot_ws.Range("A1").Font.Bold = True

        last_row = source_ws.Cells(source_ws.Rows.Count, 1).End(-4162).Row
        last_col = source_ws.Cells(1, source_ws.Columns.Count).End(-4159).Column
        source_range = source_ws.Range(source_ws.Cells(1, 1), source_ws.Cells(last_row, last_col))
        source_address = f"'{source_sheet}'!{source_range.GetAddress(RowAbsolute=True, ColumnAbsolute=True, ReferenceStyle=1)}"

        pivot_cache = workbook.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=source_address)
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=pivot_ws.Range("A3"),
            TableName=pivot_name,
        )

        pivot_table.PivotFields("Category").Orientation = XL_ROW_FIELD
        pivot_table.PivotFields("Category").Position = 1
        pivot_table.PivotFields("Subcategory").Orientation = XL_ROW_FIELD
        pivot_table.PivotFields("Subcategory").Position = 2
        pivot_table.PivotFields("Subcategory").Subtotals = (False,) * 12

        data_field = pivot_table.AddDataField(pivot_table.PivotFields("Total"), "Sum of Total", XL_SUM)
        data_field.NumberFormat = ACCOUNTING_IDR_FORMAT

        pivot_table.RowAxisLayout(XL_COMPACT_ROW)
        pivot_table.ShowDrillIndicators = True
        pivot_table.InGridDropZones = False
        pivot_table.DisplayFieldCaptions = True
        pivot_table.HasAutoFormat = True
        # Excel kadang tidak menampilkan grand total bawah di compact layout
        # bila hanya salah satu flag total yang aktif.
        pivot_table.ColumnGrand = True
        pivot_table.RowGrand = True
        pivot_table.GrandTotalName = "Grand Total"
        pivot_table.ShowTableStyleRowStripes = False
        pivot_table.CompactLayoutRowHeader = "Row Labels"
        pivot_table.RepeatAllLabels(2)

        try:
            pivot_table.DataBodyRange.NumberFormat = ACCOUNTING_IDR_FORMAT
            pivot_table.DataBodyRange.NumberFormatLocal = ACCOUNTING_IDR_FORMAT_LOCAL
        except Exception:
            pass

        try:
            pivot_table.TableRange1.Columns(2).NumberFormat = ACCOUNTING_IDR_FORMAT
            pivot_table.TableRange1.Columns(2).NumberFormatLocal = ACCOUNTING_IDR_FORMAT_LOCAL
        except Exception:
            pass

        pivot_ws.Columns("A:B").AutoFit()

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a real Excel PivotTable from a source sheet.")
    parser.add_argument("--source", required=True, type=Path, help="Source workbook path.")
    parser.add_argument("--source-sheet", required=True, help="Source worksheet name.")
    parser.add_argument("--output", required=True, type=Path, help="Output workbook path.")
    parser.add_argument("--pivot-sheet", default="Pivot_Category_Subcategory", help="Pivot worksheet name.")
    parser.add_argument("--pivot-name", help="Pivot table object name. Optional; auto-generated from pivot sheet name.")
    args = parser.parse_args()

    pivot_name = args.pivot_name
    if not pivot_name:
        parts = re.findall(r"[A-Za-z0-9]+", args.pivot_sheet)
        pivot_name = "".join(part[:1].upper() + part[1:] for part in parts) or "PivotTableAuto"

    build_pivot(
        source_path=args.source,
        source_sheet=args.source_sheet,
        output_path=args.output,
        pivot_sheet=args.pivot_sheet,
        pivot_name=pivot_name,
    )
    print(f"Pivot created in '{args.output}' on sheet '{args.pivot_sheet}'.")


if __name__ == "__main__":
    main()
