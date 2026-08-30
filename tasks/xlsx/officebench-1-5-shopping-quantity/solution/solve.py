#!/usr/bin/env python3
"""OfficeBench 1-5/1 的逐题参考解法。"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from openpyxl import load_workbook

INPUT_NAME = "shopping_list_input.xlsx"
OUTPUT_NAME = "shopping_list.xlsx"
EXPECTED_INPUT_SHA256 = "e6468e5452830bc4bd8767fb9a95c0ea95a65d876e306581437889e43c2a821a"


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
        raise ValueError("shopping_list.xlsx 输入哈希不匹配")
    workbook = load_workbook(source)
    sheet = workbook["Sheet1"]
    if sheet["A2"].value != "apple" or sheet["B2"].value != 2:
        raise ValueError("输入购物清单不在预期状态")
    sheet["B2"] = sheet["B2"].value / 2
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    args.output_dir.mkdir(parents=True, exist_ok=True)
    workbook.save(args.output_dir / OUTPUT_NAME)
    workbook.close()


if __name__ == "__main__":
    main()
