#!/usr/bin/env python3
"""OfficeBench 1-11/0 的逐题参考解法。"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from openpyxl import load_workbook

INPUT_NAME = "score_input.xlsx"
OUTPUT_NAME = "sorted_score.xlsx"
EXPECTED_INPUT_SHA256 = "b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


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
    rows = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row))
    rows.sort(key=lambda row: row[1].value, reverse=True)
    values = [[cell.value for cell in row] for row in rows]
    for row_index, row_values in enumerate(values, start=2):
        for col_index, value in enumerate(row_values, start=1):
            sheet.cell(row_index, col_index).value = value
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    args.output_dir.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output_dir / OUTPUT_NAME)
    workbook.close()


if __name__ == "__main__":
    main()
