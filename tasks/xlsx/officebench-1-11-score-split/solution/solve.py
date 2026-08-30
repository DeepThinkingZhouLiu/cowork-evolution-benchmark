#!/usr/bin/env python3
"""OfficeBench 1-11/1 的逐题参考解法。"""

from __future__ import annotations

import argparse
import hashlib
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook

INPUT_NAME = "score_input.xlsx"
EXPECTED_INPUT_SHA256 = "b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def write_column(source_sheet, output_dir: Path, name: str, column: int) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = source_sheet.title
    for row in range(1, source_sheet.max_row + 1):
        for target_col, source_col in enumerate((1, column), start=1):
            source_cell = source_sheet.cell(row, source_col)
            target_cell = sheet.cell(row, target_col, source_cell.value)
            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.number_format = source_cell.number_format
            if source_cell.number_format:
                target_cell.number_format = source_cell.number_format
    sheet.column_dimensions["A"].width = source_sheet.column_dimensions["A"].width
    sheet.column_dimensions["B"].width = source_sheet.column_dimensions["B"].width
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(output_dir / f"{name}.xlsx")
    workbook.close()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    source = args.input_dir / INPUT_NAME
    if sha256(source) != EXPECTED_INPUT_SHA256:
        raise ValueError("score.xlsx 输入哈希不匹配")
    workbook = load_workbook(source)
    sheet = workbook["Sheet1"]
    if sheet.max_column < 3:
        raise ValueError("输入缺少两项成绩列")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_column(sheet, args.output_dir, "midterm1", 2)
    write_column(sheet, args.output_dir, "midterm2", 3)
    workbook.close()


if __name__ == "__main__":
    main()
