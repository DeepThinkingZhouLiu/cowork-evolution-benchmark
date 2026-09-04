#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-2-19-teaching-report"
WEIGHTS = {"R001": 0.15, "R002": 0.25, "R003": 0.18, "R004": 0.18, "R005": 0.14, "R006": 0.10}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def band(score: float) -> str:
    return "A" if score >= 90 else "B" if score >= 80 else "C" if score >= 70 else "D" if score >= 60 else "F"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    source = Path(__file__).parents[1] / "data/input_files/score.xlsx"
    output = args.artifact_dir / "teaching.docx"
    results = []
    valid = output.is_file() and zipfile.is_zipfile(output)
    text = ""
    tables = []
    try:
        document = Document(output) if valid else None
        if document is not None:
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            tables = [[[cell.text.strip() for cell in row.cells] for row in table.rows] for table in document.tables]
        valid = valid and bool(text.strip())
    except Exception as error:
        valid = False
        text = str(error)
    results.append(criterion("R001", valid, f"可解析的原生 OOXML DOCX={valid}"))
    try:
        workbook = load_workbook(source, data_only=True, read_only=True)
        rows = list(workbook.active.iter_rows(min_row=2, max_col=2, values_only=True))
        workbook.close()
        records = [(str(name).strip(), float(score)) for name, score in rows if name is not None]
        grouped = defaultdict(list)
        for name, score in records:
            grouped[name].append(score)
        distribution = Counter(band(score) for _, score in records)
        average = sum(score for _, score in records) / len(records)
        overview_ok = all(token in text for token in [f"Grade records: {len(records)}", f"Students: {len(grouped)}", f"Class average: {average:.2f}"])
        distribution_ok = all(re.search(rf"\b{key}\s*=\s*{value}\b", text) for key, value in distribution.items())
        results.append(criterion("R002", overview_ok, f"记录={len(records)}，学生={len(grouped)}，平均={average:.2f}"))
        results.append(criterion("R003", distribution_ok, f"分布={dict(distribution)}"))
        summary = {row[0]: row[1:] for row in tables[0][1:]} if tables and tables[0] and len(tables[0][0]) == 5 else {}
        summary_ok = len(summary) == len(grouped) and all(
            name in summary
            and int(float(summary[name][0])) == len(scores)
            and abs(float(summary[name][1]) - sum(scores) / len(scores)) < 0.011
            and abs(float(summary[name][2]) - max(scores)) < 0.01
            and abs(float(summary[name][3]) - min(scores)) < 0.01
            for name, scores in grouped.items()
        )
        results.append(criterion("R004", summary_ok, "逐学生记录数、平均、最高和最低分独立重算"))
        support = [name for name, scores in grouped.items() if sum(scores) / len(scores) < 70]
        praise = [name for name, scores in grouped.items() if sum(scores) / len(scores) >= 90]
        priorities_ok = "Priority support:" in text and "High-performing students:" in text and all(name in text for name in support + praise)
        results.append(criterion("R005", priorities_ok, f"辅导={support}，表扬={praise}"))
        digest = hashlib.sha256(source.read_bytes()).hexdigest()
        results.append(criterion("R006", digest == "e151e7aff861749d1fce7bea0e584263cdc5182df25f9b6a9a96cb62ca9b19f8", f"输入 SHA256={digest}"))
    except Exception as error:
        for identifier in ("R002", "R003", "R004", "R005", "R006"):
            results.append(criterion(identifier, False, f"输入驱动检查失败：{error}"))
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[key] * values.get(key, 0.0) for key in WEIGHTS)
    passed = reward >= 0.8 and values.get("R001") == 1.0 and values.get("R002") == 1.0 and values.get("R004") == 1.0
    payload = {"task_id": TASK_ID, "rubric_version": "officebench-2-19-teaching-report-v0.2", "judge_version": "officebench-2-19-teaching-report-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [], "criterion_results": results, "reward": round(reward, 6), "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
