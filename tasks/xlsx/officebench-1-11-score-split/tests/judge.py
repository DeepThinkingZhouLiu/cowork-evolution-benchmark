#!/usr/bin/env python3
"""从输入独立重算两个考试工作表，检查单一 XLSX 交付。"""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from openpyxl import load_workbook

TASK_ID = "cowork-evo/officebench-1-11-score-split"
RUBRIC_VERSION = "officebench-1-11-score-split-v0.2"
JUDGE_VERSION = "officebench-1-11-score-split-deterministic-v0.2"
INPUT_NAME = "score_input.xlsx"
OUTPUT_NAME = "midterm_split.xlsx"
EXPECTED_INPUT_SHA256 = "b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973"
WEIGHTS = {"R001": 0.20, "R002": 0.30, "R003": 0.30, "R004": 0.10, "R005": 0.05, "R006": 0.05}


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def criterion(cid: str, score: bool | float, evidence: str) -> dict:
    return {"criterion_id": cid, "score": float(score), "evidence": evidence, "method": "deterministic"}


def style_signature(cell) -> tuple[str, ...]:
    return tuple(str(getattr(cell, attr)) for attr in ("font", "fill", "border", "alignment", "protection", "number_format"))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    root = Path(__file__).parents[1]
    input_path = root / "data/input_files" / INPUT_NAME
    output_path = args.artifact_dir / OUTPUT_NAME
    results: list[dict] = []
    source = output = None
    try:
        source = load_workbook(input_path, data_only=False)
        source_sheet = source["Sheet1"]
        valid_file = output_path.is_file() and zipfile.is_zipfile(output_path)
        if not valid_file:
            results.append(criterion("R001", False, "midterm_split.xlsx 缺失或不是有效 OOXML"))
            for cid in ("R002", "R003", "R004", "R005", "R006"):
                results.append(criterion(cid, False, "缺少有效交付，无法检查"))
        else:
            output = load_workbook(output_path, data_only=False)
            structure_ok = output.sheetnames == ["midterm1", "midterm2"]
            results.append(criterion("R001", structure_ok, f"工作表顺序={output.sheetnames}"))
            expected = {"midterm1": ("midterm1", 2), "midterm2": ("midterm2", 3)}
            values_ok = {}
            style_ok = True
            for sheet_name, (header, source_col) in expected.items():
                if sheet_name not in output.sheetnames:
                    values_ok[sheet_name] = False
                    style_ok = False
                    continue
                sheet = output[sheet_name]
                values = [[sheet.cell(row, col).value for col in (1, 2)] for row in range(1, sheet.max_row + 1)]
                expected_values = [[source_sheet.cell(row, 1).value, source_sheet.cell(row, source_col).value] for row in range(1, source_sheet.max_row + 1)]
                values_ok[sheet_name] = values == expected_values and values[:1] == [["Name", header]] and sheet.max_column == 2 and sheet.max_row == source_sheet.max_row
                for row in range(1, source_sheet.max_row + 1):
                    for out_col, source_column in ((1, 1), (2, source_col)):
                        style_ok &= style_signature(sheet.cell(row, out_col)) == style_signature(source_sheet.cell(row, source_column))
                score_letter = "B" if source_col == 2 else "C"
                style_ok &= sheet.column_dimensions["A"].width == source_sheet.column_dimensions["A"].width
                style_ok &= sheet.column_dimensions["B"].width == source_sheet.column_dimensions[score_letter].width
            results.append(criterion("R002", values_ok.get("midterm1", False), "midterm1 工作表由输入 Name+midterm1 独立投影"))
            results.append(criterion("R003", values_ok.get("midterm2", False), "midterm2 工作表由输入 Name+midterm2 独立投影"))
            results.append(criterion("R004", structure_ok and style_ok, f"行列结构和单元格样式保持={style_ok}"))
            source_hash_ok = digest(input_path) == EXPECTED_INPUT_SHA256
            results.append(criterion("R005", source_hash_ok, f"输入 SHA256={digest(input_path)}"))
            editable = all(output[name].protection.sheet is not True for name in output.sheetnames) if structure_ok else False
            results.append(criterion("R006", editable, "两个工作表均未保护且为原生单元格"))
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
        results.extend(criterion(cid, False, "Judge 读取失败") for cid in ("R002", "R003", "R004", "R005", "R006"))
    finally:
        if source:
            source.close()
        if output:
            output.close()
    by_id = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[cid] * by_id.get(cid, 0.0) for cid in WEIGHTS)
    passed = by_id.get("R001", 0.0) == 1.0 and by_id.get("R002", 0.0) == 1.0 and by_id.get("R003", 0.0) == 1.0 and reward >= 0.7
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest(output_path)] if output_path.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
