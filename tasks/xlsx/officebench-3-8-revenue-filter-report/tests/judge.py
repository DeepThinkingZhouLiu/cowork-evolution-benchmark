#!/usr/bin/env python3
"""确定性 Judge：从收入输入重算筛选集合并核对 Word 报告。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

TASK_ID = "cowork-evo/officebench-3-8-revenue-filter-report"
RUBRIC_VERSION = "officebench-3-8-revenue-filter-report-v0.1"
JUDGE_VERSION = "officebench-3-8-revenue-filter-report-deterministic-v0.1"
OUTPUTS = ["new.xlsx", "report.docx"]
THRESHOLD = 2_000_000
EXPECTED_SHA256 = "b7a5d913221fafb9f545e9367df2bcf07e665b93baefe726788a88532c937e28"
WEIGHTS = {"R001": 0.15, "R002": 0.30, "R003": 0.10, "R004": 0.15, "R005": 0.10, "R006": 0.10, "R007": 0.05, "R008": 0.05}


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def criterion(cid: str, score: bool | float, evidence: str) -> dict:
    return {"criterion_id": cid, "score": float(score), "evidence": evidence, "method": "deterministic"}


def document_text(path: Path) -> tuple[str, list[str], list[list[list[str]]]]:
    document = Document(path)
    parts = [paragraph.text for paragraph in document.paragraphs]
    headings = [paragraph.text for paragraph in document.paragraphs if paragraph.style.name.startswith("Heading") or paragraph.style.name == "Title"]
    tables = []
    for table in document.tables:
        table_rows = []
        for row in table.rows:
            values = [cell.text for cell in row.cells]
            parts.extend(values)
            table_rows.append(values)
        tables.append(table_rows)
    return "\n".join(parts), headings, tables


def normalized_digits(text: str) -> str:
    return re.sub(r"[^0-9]", "", text)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    source_path = Path(__file__).parents[1] / "data/input_files/company_revenues.xlsx"
    results: list[dict] = []
    try:
        missing = [name for name in OUTPUTS if not (args.artifact_dir / name).is_file()]
        native = not missing and all(zipfile.is_zipfile(args.artifact_dir / name) for name in OUTPUTS)
        results.append(criterion("R001", native, f"缺失={missing}; OOXML={native}"))
        source_wb = load_workbook(source_path, data_only=True)
        source_rows = [(int(year), int(revenue)) for year, revenue in source_wb["Sheet1"].iter_rows(min_row=2, values_only=True)]
        source_wb.close()
        expected = [row for row in source_rows if row[1] >= THRESHOLD]
        removed = [row for row in source_rows if row[1] < THRESHOLD]
        candidate_wb = load_workbook(args.artifact_dir / "new.xlsx", data_only=True)
        candidate_sheet = candidate_wb["Sheet1"] if "Sheet1" in candidate_wb.sheetnames else candidate_wb.active
        headers = tuple(cell.value for cell in candidate_sheet[1][:2])
        candidate_rows = [(int(year), int(revenue)) for year, revenue in candidate_sheet.iter_rows(min_row=2, max_col=2, values_only=True) if year is not None]
        results.append(criterion("R002", candidate_rows == expected, f"期望 17 行；输出 {len(candidate_rows)} 行"))
        invalid = [row for row in candidate_rows if row[1] < THRESHOLD]
        set_ok = len(candidate_rows) == len(set(candidate_rows)) and not invalid and set(candidate_rows) == set(expected)
        results.append(criterion("R003", set_ok, f"低于阈值={invalid}; 唯一行={len(set(candidate_rows))}"))
        text, headings, tables = document_text(args.artifact_dir / "report.docx")
        compact = normalized_digits(text)
        counts_ok = "2000000" in compact and re.search(r"(?:input|before|输入|筛选前)[^0-9]{0,20}20", text, re.I) and re.search(r"(?:retain|after|保留|筛选后)[^0-9]{0,20}17", text, re.I) and re.search(r"(?:remove|delete|删除)[^0-9]{0,20}3", text, re.I)
        results.append(criterion("R004", bool(counts_ok), "阈值及 20/17/3 摘要检查"))
        reported_removed = set()
        for table in tables:
            for row in table[1:]:
                if len(row) < 2:
                    continue
                year_digits = re.sub(r"[^0-9]", "", row[0])
                revenue_digits = re.sub(r"[^0-9]", "", row[1])
                if len(year_digits) == 4 and revenue_digits:
                    reported_removed.add((int(year_digits), int(revenue_digits)))
        removed_ok = reported_removed == set(removed)
        results.append(criterion("R005", removed_ok, f"应删除={removed}"))
        cross_ok = headers == ("Year", "Revenue ($)") and candidate_rows == expected and removed_ok and bool(counts_ok)
        results.append(criterion("R006", cross_ok, f"表头={headers}; 报告与 Excel 一致={cross_ok}"))
        professional = candidate_sheet.auto_filter.ref is not None and candidate_sheet.freeze_panes == "A2" and len(headings) >= 2 and len(tables) >= 1
        results.append(criterion("R007", professional, f"筛选={candidate_sheet.auto_filter.ref}; 标题数={len(headings)}; 表格数={len(tables)}"))
        results.append(criterion("R008", digest(source_path) == EXPECTED_SHA256, f"输入 SHA256={digest(source_path)}"))
        candidate_wb.close()
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
    by_id = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[cid] * by_id.get(cid, 0.0) for cid in WEIGHTS)
    passed = reward >= 0.8 and all(by_id.get(cid) == 1.0 for cid in ("R001", "R002", "R003", "R004", "R006"))
    artifacts = [digest(args.artifact_dir / name) for name in OUTPUTS if (args.artifact_dir / name).is_file()]
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": artifacts, "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
