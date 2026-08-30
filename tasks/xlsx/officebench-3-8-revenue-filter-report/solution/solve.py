#!/usr/bin/env python3
"""OfficeBench 3-8/0 的逐题参考解法。"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from openpyxl import load_workbook

INPUT_NAME = "company_revenues.xlsx"
EXPECTED_SHA256 = "b7a5d913221fafb9f545e9367df2bcf07e665b93baefe726788a88532c937e28"
THRESHOLD = 2_000_000


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    source = args.input_dir / INPUT_NAME
    if digest(source) != EXPECTED_SHA256:
        raise ValueError("收入输入哈希不匹配")
    workbook = load_workbook(source)
    sheet = workbook["Sheet1"]
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    removed = [(int(year), int(revenue)) for year, revenue in rows if revenue < THRESHOLD]
    for row_number in range(sheet.max_row, 1, -1):
        if sheet.cell(row_number, 2).value < THRESHOLD:
            sheet.delete_rows(row_number)
    sheet.auto_filter.ref = f"A1:B{sheet.max_row}"
    sheet.freeze_panes = "A2"
    sheet.column_dimensions["A"].width = 12
    sheet.column_dimensions["B"].width = 18
    for cell in sheet["B"][1:]:
        cell.number_format = "$#,##0"
    args.output_dir.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output_dir / "new.xlsx")
    workbook.close()

    document = Document()
    document.add_heading("Company Revenue Filter Report", 0)
    document.add_heading("Executive Summary", level=1)
    document.add_paragraph("Annual revenue records below $2,000,000 were removed. Records equal to the threshold are retained.")
    document.add_paragraph(f"Input rows: {len(rows)}. Retained rows: {len(rows) - len(removed)}. Removed rows: {len(removed)}.")
    document.add_heading("Removed Years", level=1)
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.rows[0].cells[0].text = "Year"
    table.rows[0].cells[1].text = "Revenue ($)"
    for year, revenue in removed:
        cells = table.add_row().cells
        cells[0].text = str(year)
        cells[1].text = f"${revenue:,.0f}"
    document.add_heading("Method", level=1)
    document.add_paragraph("The source order and original revenue values were preserved for all retained years in new.xlsx.")
    document.save(args.output_dir / "report.docx")


if __name__ == "__main__":
    main()
