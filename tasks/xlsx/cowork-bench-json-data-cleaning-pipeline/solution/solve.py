#!/usr/bin/env python3
"""生成数据清洗 JSON、三表 Excel、离线 Dashboard 与进度记录。"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


SCRIPT_NAME = "clean_data.py"
SALES_NAME = "cleaned_sales.json"
INVENTORY_NAME = "cleaned_inventory.json"
ANALYSIS_NAME = "pipeline_results.json"
REPORT_NAME = "Data_Pipeline_Report.xlsx"
DASHBOARD_NAME = "Pipeline_Dashboard.xlsx"
NOTE_NAME = "progress_note.json"
NAVY, BLUE, PALE, WHITE, INK = "17365D", "2F75B5", "EAF2F8", "FFFFFF", "1F2937"


def add_table(sheet, ref: str, name: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)


def style_sheet(sheet, widths: dict[str, float], table_ref: str, table_name: str) -> None:
    header_fill = PatternFill("solid", fgColor=NAVY)
    line = Side(style="thin", color="D7E0EA")
    for cell in sheet[1]:
        cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=10, color=INK)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=line)
        sheet.row_dimensions[row[0].row].height = 24
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = table_ref
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.print_options.horizontalCentered = True
    sheet.page_margins.left = 0.3
    sheet.page_margins.right = 0.3
    sheet.page_margins.top = 0.4
    sheet.page_margins.bottom = 0.4
    sheet.oddFooter.center.text = "Cowork-Bench · Reproducible Data Cleaning"
    add_table(sheet, table_ref, table_name)


def write_report(path: Path, result: dict) -> None:
    workbook = Workbook()
    workbook.properties.creator = "Cowork Evolution Benchmark"
    workbook.properties.title = "Data Pipeline Report"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"

    sales = workbook.active
    sales.title = "Sales Analysis"
    sales_headers = ["Product_Name", "Total_Quantity_Sold", "Total_Revenue", "Avg_Unit_Price"]
    sales.append(sales_headers)
    for row in result["sales_analysis"]:
        sales.append([row[key] for key in sales_headers])
    for cell in sales["C"][1:] + sales["D"][1:]:
        cell.number_format = '₽#,##0.00'
    style_sheet(sales, {"A": 28, "B": 23, "C": 20, "D": 20}, "A1:D6", "Pipeline_Sales_Analysis")
    for row in sales.iter_rows(min_row=2, min_col=2, max_col=4):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    inventory = workbook.create_sheet("Inventory Status")
    inventory.append(["Product_Name", "Current_Stock", "Reorder_Point", "Status"])
    for index, row in enumerate(result["inventory_status"], start=2):
        inventory.append([row["Product_Name"], row["Current_Stock"], row["Reorder_Point"], None])
        inventory[f"D{index}"] = f'=IF(B{index}<=C{index},"LOW","OK")'
    style_sheet(inventory, {"A": 28, "B": 20, "C": 20, "D": 16}, "A1:D6", "Pipeline_Inventory_Status")
    for row in inventory.iter_rows(min_row=2, min_col=2, max_col=4):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    combined = workbook.create_sheet("Combined View")
    combined_headers = ["Product_Name", "Total_Sold", "Current_Stock", "Revenue", "Stock_Status"]
    combined.append(combined_headers)
    for row in result["combined_view"]:
        combined.append([row[key] for key in combined_headers])
    for cell in combined["D"][1:]:
        cell.number_format = '₽#,##0.00'
    style_sheet(combined, {"A": 28, "B": 17, "C": 20, "D": 18, "E": 18}, "A1:E6", "Pipeline_Combined_View")
    for row in combined.iter_rows(min_row=2, min_col=2, max_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    workbook.save(path)


def write_dashboard(path: Path, result: dict) -> None:
    workbook = Workbook()
    workbook.properties.creator = "Cowork Evolution Benchmark"
    workbook.properties.title = "Pipeline Dashboard · Offline Google Sheet Replacement"
    sheet = workbook.active
    sheet.title = "Pipeline Dashboard"
    headers = ["Product_Name", "Total_Sold", "Current_Stock", "Revenue", "Stock_Status"]
    sheet.append(headers)
    for row in result["combined_view"]:
        sheet.append([row[key] for key in headers])
    for cell in sheet["D"][1:]:
        cell.number_format = '₽#,##0.00'
    style_sheet(sheet, {"A": 28, "B": 17, "C": 20, "D": 18, "E": 18}, "A1:E6", "Offline_Pipeline_Dashboard")
    for row in sheet.iter_rows(min_row=2, min_col=2, max_col=5):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    workbook.save(path)


def write_progress(path: Path, result: dict) -> None:
    summary = result["summary"]
    payload = {
        "entities": [
            {
                "name": "Data_Pipeline",
                "entityType": "project",
                "observations": [
                    "Read sales and inventory API response JSON files from workspace",
                    "Created clean_data.py to keep the first order_id/product_id and normalize only product_name",
                    f"Cleaned data: {summary['unique_sales_count']} unique sales records and {summary['unique_inventory_count']} unique inventory products",
                    f"Removed duplicates: {summary['sales_duplicates_removed']} sales and {summary['inventory_duplicates_removed']} inventory records",
                    "Created Data_Pipeline_Report.xlsx with Sales Analysis, Inventory Status, and Combined View sheets",
                    "Created offline Pipeline_Dashboard.xlsx with the same data as Combined View",
                ],
            }
        ],
        "relations": [],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def solve(input_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    source = Path(__file__).with_name(SCRIPT_NAME)
    script = output_dir / SCRIPT_NAME
    if source.resolve() != script.resolve():
        shutil.copy2(source, script)
    subprocess.run(
        [
            sys.executable,
            str(script),
            "--sales", str(input_dir / "sales_api_response.json"),
            "--inventory", str(input_dir / "inventory_api_response.json"),
            "--sales-output", str(output_dir / SALES_NAME),
            "--inventory-output", str(output_dir / INVENTORY_NAME),
            "--analysis-output", str(output_dir / ANALYSIS_NAME),
        ],
        check=True,
        timeout=120,
    )
    result = json.loads((output_dir / ANALYSIS_NAME).read_text(encoding="utf-8"))
    write_report(output_dir / REPORT_NAME, result)
    write_dashboard(output_dir / DASHBOARD_NAME, result)
    write_progress(output_dir / NOTE_NAME, result)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
