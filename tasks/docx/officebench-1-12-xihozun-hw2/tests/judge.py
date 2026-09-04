#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

TASK_ID = 'cowork-evo/officebench-1-12-xihozun-hw2'
OUTPUT_NAME = 'xihozun_hw2.docx'
INPUT_NAME = 'CS161_Class_Roster.xlsx'
INPUT_SHA = '19f2b63d8ce3010d55b28058226b0a4db2ee38e5567ab1f9ef5ae4dedad7092a'
KIND = 'xihozun_hw2'
WEIGHTS = {"R001": 0.20, "R002": 0.25, "R003": 0.18, "R004": 0.14, "R005": 0.13, "R006": 0.10}


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def source_rows(path: Path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    values = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() if value is not None else "" for value in values[0]]
    return headers, values[1:]


def normalized_course(value: object) -> str:
    return "".join(str(value or "").upper().split())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    source = Path(__file__).parents[1] / "data/input_files" / INPUT_NAME
    output = args.artifact_dir / OUTPUT_NAME
    results = []
    valid = output.is_file() and zipfile.is_zipfile(output)
    text = ""
    paragraphs = []
    try:
        if valid:
            document = Document(output)
            paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
            text = "\n".join(paragraphs)
            valid = bool(text)
    except Exception as error:
        valid = False
        text = f"read error: {error}"
    results.append(criterion("R001", valid, f"DOCX readable={valid}, paragraphs={len(paragraphs)}"))
    try:
        headers, data = source_rows(source)
        index = {name: position for position, name in enumerate(headers)}
        if KIND == "roster_counts":
            expected_total = len(data)
            expected_seniors = sum(1 for row in data if str(row[index["Academic Standing"]]).strip() == "Senior")
            total_ok = bool(re.search(rf"Total students\s*:\s*{expected_total}\b", text))
            senior_ok = bool(re.search(rf"Senior students\s*:\s*{expected_seniors}\b", text))
            core_ok = total_ok and senior_ok
            detail = f"expected total={expected_total}, seniors={expected_seniors}; matched={total_ok and senior_ok}"
            structure_ok = len(paragraphs) == 3 and paragraphs[0] == "CS161 Roster Counts"
        elif KIND == "final_pass_count":
            expected = sum(1 for row in data if float(row[index["Final"]]) >= 60)
            core_ok = bool(re.search(rf"Students with final score\s*>=\s*60\s*:\s*{expected}\b", text))
            detail = f"expected final>=60={expected}; matched={core_ok}"
            structure_ok = len(paragraphs) == 2 and paragraphs[0] == "CS161 Final Score Count"
        elif KIND == "xihozun_hw2":
            matches = [row for row in data if str(row[index["Last Name"]]).strip() == "Xihozun"]
            expected = matches[0][index["HW2"]] if len(matches) == 1 else None
            core_ok = bool(expected is not None and re.search(rf"Xihozun HW2 score\s*:\s*{re.escape(str(expected))}\b", text))
            detail = f"expected Xihozun HW2={expected}; matched={core_ok}"
            structure_ok = len(paragraphs) == 2 and paragraphs[0] == "Xihozun HW2 Lookup"
        elif KIND == "prerequisite_check":
            found = {normalized_course(row[index["Course ID"]]) for row in data}
            required = ["CS101", "CS102", "MATH151", "MATH152"]
            expected = all(course in found for course in required)
            verdict = "Yes" if expected else "No"
            core_ok = bool(re.search(rf"CS161 prerequisites satisfied\s*:\s*{verdict}\b", text))
            evidence_ok = all(course in normalized_course(text) for course in required)
            detail = f"expected={verdict}; required evidence={evidence_ok}"
            structure_ok = len(paragraphs) == 3 and paragraphs[0] == "CS161 Prerequisite Check"
            core_ok = core_ok and evidence_ok
        elif KIND == "transfer_eligibility":
            cs_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("CS"))
            verdict = "Yes" if cs_count >= 5 else "No"
            count_ok = bool(re.search(rf"CS courses completed\s*:\s*{cs_count}\b", text))
            verdict_ok = bool(re.search(rf"Eligible for CS transfer\s*:\s*{verdict}\b", text))
            core_ok = count_ok and verdict_ok
            detail = f"expected CS count={cs_count}, verdict={verdict}; matched={core_ok}"
            structure_ok = len(paragraphs) == 3 and paragraphs[0] == "CS Transfer Eligibility"
        elif KIND == "graduate_requirement":
            cs_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("CS"))
            math_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("MATH"))
            verdict = "Yes" if cs_count >= 8 and math_count >= 5 else "No"
            count_ok = bool(re.search(rf"CS courses completed\s*:\s*{cs_count}\b", text)) and bool(re.search(rf"MATH courses completed\s*:\s*{math_count}\b", text))
            verdict_ok = bool(re.search(rf"Graduate requirement satisfied\s*:\s*{verdict}\b", text))
            core_ok = count_ok and verdict_ok
            detail = f"expected CS={cs_count}, MATH={math_count}, verdict={verdict}; matched={core_ok}"
            structure_ok = len(paragraphs) == 4 and paragraphs[0] == "Graduate Course Requirement"
        else:
            raise ValueError(KIND)
        results.append(criterion("R002", core_ok, detail))
        results.append(criterion("R003", core_ok, "all input-derived target facts are present"))
        results.append(criterion("R004", structure_ok, f"expected editable heading/result structure={structure_ok}"))
        results.append(criterion("R005", valid and not any(token in text.lower() for token in ("todo", "placeholder", "nan")), "no placeholder or invalid value"))
        actual_sha = hashlib.sha256(source.read_bytes()).hexdigest()
        results.append(criterion("R006", actual_sha == INPUT_SHA, f"input SHA256={actual_sha}"))
    except Exception as error:
        results.extend([
            criterion("R002", False, f"input-driven check failed: {error}"),
            criterion("R003", False, "not executed"),
            criterion("R004", False, "not executed"),
            criterion("R005", False, "not executed"),
            criterion("R006", False, "not executed"),
        ])
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[key] * values.get(key, 0.0) for key in WEIGHTS)
    passed = values.get("R001", 0.0) == 1.0 and values.get("R002", 0.0) == 1.0 and reward >= 0.8
    payload = {
        "task_id": TASK_ID,
        "rubric_version": f"{TASK_ID.rsplit('/', 1)[-1]}-v0.1",
        "judge_version": f"{TASK_ID.rsplit('/', 1)[-1]}-deterministic-v0.1",
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": round(reward, 6),
        "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
