#!/usr/bin/env python3
"""确定性 Judge：从输入独立投影两个考试列。"""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from openpyxl import load_workbook

TASK_ID = "cowork-evo/officebench-1-11-score-split"
RUBRIC_VERSION = "officebench-1-11-score-split-v0.1"
JUDGE_VERSION = "officebench-1-11-score-split-deterministic-v0.1"
INPUT_NAME = "score_input.xlsx"
EXPECTED_INPUT_SHA256 = "b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973"
WEIGHTS = {"R001": 0.20, "R002": 0.30, "R003": 0.30, "R004": 0.10, "R005": 0.05, "R006": 0.05}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def criterion(cid: str, score: bool | float, evidence: str) -> dict:
    return {"criterion_id": cid, "score": float(score), "evidence": evidence, "method": "deterministic"}


def read_output(path: Path):
    if not path.is_file() or not zipfile.is_zipfile(path):
        raise ValueError(f"{path.name} 缺失或不是有效 OOXML")
    return load_workbook(path, data_only=False)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    input_path = Path(__file__).parents[1] / "data/input_files" / INPUT_NAME
    paths = {name: args.artifact_dir / name for name in ("midterm1.xlsx", "midterm2.xlsx")}
    results: list[dict] = []
    source = None
    outputs = []
    try:
        source = load_workbook(input_path, data_only=False)
        sws = source["Sheet1"]
        outputs = [read_output(paths[name]) for name in paths]
        expected = {
            "midterm1.xlsx": [[sws.cell(row, 1).value, sws.cell(row, 2).value] for row in range(1, sws.max_row + 1)],
            "midterm2.xlsx": [[sws.cell(row, 1).value, sws.cell(row, 3).value] for row in range(1, sws.max_row + 1)],
        }
        actual = {}
        for name, workbook in zip(paths, outputs):
            ws = workbook["Sheet1"]
            actual[name] = [[cell.value for cell in row] for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=2)]
        structure_ok = all(workbook.sheetnames == source.sheetnames and workbook["Sheet1"].max_row == sws.max_row and workbook["Sheet1"].max_column == 2 for workbook in outputs)
        files_ok = all(paths[name].is_file() for name in paths)
        results.append(criterion("R001", files_ok and structure_ok, f"文件齐全={files_ok}; 结构保持={structure_ok}"))
        r2 = actual["midterm1.xlsx"] == expected["midterm1.xlsx"] and actual["midterm1.xlsx"][0] == ["Name", "midterm1"]
        r3 = actual["midterm2.xlsx"] == expected["midterm2.xlsx"] and actual["midterm2.xlsx"][0] == ["Name", "midterm2"]
        results.append(criterion("R002", r2, f"midterm1 独立投影一致={r2}"))
        results.append(criterion("R003", r3, f"midterm2 独立投影一致={r3}"))
        style_ok = True
        for name, workbook, source_col in (("midterm1.xlsx", outputs[0], 2), ("midterm2.xlsx", outputs[1], 3)):
            ws = workbook["Sheet1"]
            for row in range(1, sws.max_row + 1):
                for out_col, src_col in ((1, 1), (2, source_col)):
                    # Separate workbooks have independent style tables; compare
                    # serialized style properties rather than internal IDs.
                    if any(str(getattr(ws.cell(row, out_col), attr)) != str(getattr(sws.cell(row, src_col), attr)) for attr in ("font", "fill", "border", "alignment", "protection", "number_format")):
                        style_ok = False
        results.append(criterion("R004", style_ok, f"两个输出的投影样式保持={style_ok}"))
        results.append(criterion("R005", digest(input_path) == EXPECTED_INPUT_SHA256, f"输入 SHA256={digest(input_path)}"))
        results.append(criterion("R006", all(workbook["Sheet1"].protection.sheet is not True for workbook in outputs), "两个工作表均可编辑"))
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
    finally:
        if source:
            source.close()
        for workbook in outputs:
            workbook.close()
    by_id = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[cid] * by_id.get(cid, 0.0) for cid in WEIGHTS)
    passed = reward >= 0.7 and all(by_id.get(cid) == 1.0 for cid in ("R001", "R002", "R003"))
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest(path) for path in paths.values() if path.is_file()], "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
