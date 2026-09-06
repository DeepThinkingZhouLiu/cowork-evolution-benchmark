#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import zipfile
from collections import Counter
from email import policy
from email.parser import BytesParser
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


TASK_ROOT = Path(__file__).resolve().parents[1]
EXPECT = json.loads((TASK_ROOT / "tests" / "expectations.json").read_text(encoding="utf-8"))
RUBRIC = {
    "task_id": "cowork-evo/officebench-2-22-midterm-report",
    "rubric_version": "officebench-2-22-midterm-report-v0.2",
    "pass_threshold": 0.8,
    "criteria": [
        {"id": "R001", "weight": 10, "type": "hurdle", "score_cap_on_failure": 0},
        {"id": "R002", "weight": 8, "type": "hurdle", "score_cap_on_failure": 0.65},
        {"id": "R003", "weight": 9, "type": "hurdle", "score_cap_on_failure": 0.65},
        {"id": "R004", "weight": 9, "type": "hurdle", "score_cap_on_failure": 0.65},
        {"id": "R005", "weight": 6, "type": "hurdle", "score_cap_on_failure": 0.65},
        {"id": "R006", "weight": 5, "type": "positive"},
        {"id": "R007", "weight": 5, "type": "hurdle", "score_cap_on_failure": 0.65},
        {"id": "R008", "weight": 4, "type": "positive"},
    ],
}


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def compact(value: object) -> str:
    return re.sub(r"[\s,，。；;：:（）()'\"`]+", "", str(value or "")).casefold()


def document_content(document: Document) -> tuple[str, int, int]:
    values = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            values.extend(cell.text for cell in row.cells)
    return "\n".join(values), len([value for value in values if compact(value)]), len(document.tables)


def result(identifier: str, score: float, evidence: str, *, triggered: bool = False) -> dict:
    return {"criterion_id": identifier, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic", "raw": {"triggered": triggered}}


def anchor_ratio(text: str, anchors: list[str]) -> tuple[float, list[str]]:
    value = compact(text)
    missing = [anchor for anchor in anchors if compact(anchor) not in value]
    return ((len(anchors) - len(missing)) / len(anchors) if anchors else 1.0), missing


def parse_ics(path: Path) -> list[str]:
    return [value.strip() for value in re.findall(r"^SUMMARY:(.*)$", path.read_text(encoding="utf-8"), re.MULTILINE)]


def parse_email(path: Path) -> dict[str, str]:
    message = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    part = message.get_body(preferencelist=("plain", "html"))
    raw = part.get_content().strip() if part else ""
    match = re.search(r"(?:^|\n)Sent at:\s*([^\r\n]+)", raw)
    if not match:
        raise ValueError(f"missing Sent at in {path.name}")
    return {"sent_at": match.group(1).strip(), "sender": str(message["From"] or ""), "subject": str(message["Subject"] or ""), "body": raw[:match.start()].strip()}


def source_hash_score() -> tuple[float, str]:
    source_dir = TASK_ROOT / "data" / "input_files"
    actual = {path.name: digest(path) for path in source_dir.iterdir() if path.is_file()}
    expected = EXPECT["input_sha256"]
    return (1.0 if actual == expected else 0.0), f"input hashes match={actual == expected}; files={len(actual)}"


def evaluate_mode(document: Document | None, text: str, native: bool) -> list[dict]:
    mode = EXPECT["mode"]
    source_dir = TASK_ROOT / "data" / "input_files"
    rows: list[dict] = []
    if mode == "event_frequency":
        counts = Counter(parse_ics(source_dir / "Alice.ics"))
        ranked = counts.most_common()
        winner, count = ranked[0]
        rows.append(result("R002", 1.0 if f"Most frequent event: {winner}" in text else 0.0, f"expected winner={winner}"))
        rows.append(result("R003", 1.0 if re.search(rf"Count:\s*{count}\b", text) else 0.0, f"expected count={count}"))
        anchors = [f"Total events: {sum(counts.values())}"] + [f"{name} {value}" for name, value in ranked]
        ratio, missing = anchor_ratio(text, anchors)
        rows.append(result("R004", ratio, f"missing={missing}"))
        rows.append(result("R005", 1.0 if document and len([p for p in document.paragraphs if compact(p.text)]) >= 6 else 0.0, "labeled editable paragraphs"))
        score, evidence = source_hash_score(); rows.append(result("R006", score, evidence))
        unique = len(ranked) < 2 or ranked[0][1] > ranked[1][1]
        rows.append(result("R007", 1.0 if (not unique or "maximum is unique" in text.casefold()) else 0.0, f"unique maximum={unique}"))
        rows.append(result("R008", 1.0 if native and not re.search(r"TODO|placeholder|nan", text, re.I) else 0.0, "native and no placeholders"))
    elif mode == "midterm_report":
        workbook = load_workbook(source_dir / "score.xlsx", data_only=True, read_only=True)
        values = list(workbook.active.iter_rows(values_only=True)); workbook.close()
        headers = [str(value).strip() for value in values[0]]; index = {value: i for i, value in enumerate(headers)}
        data = [row for row in values[1:] if row[index["Name"]] is not None]
        m1 = sum(float(row[index["midterm1"]]) for row in data) / len(data)
        m2 = sum(float(row[index["midterm2"]]) for row in data) / len(data)
        checks = [("R002", f"Students\n{len(data)}", len(data)), ("R003", f"Midterm 1 average\n{m1:.2f}", m1), ("R004", f"Midterm 2 average\n{m2:.2f}", m2), ("R005", f"{m2-m1:+.2f}", m2-m1)]
        for identifier, anchor, expected in checks:
            rows.append(result(identifier, 1.0 if compact(anchor) in compact(text) else 0.0, f"expected={expected}"))
        rows.append(result("R006", 1.0 if document and len(document.tables) == 1 and len(document.tables[0].rows) >= 5 else 0.0, "native result table"))
        score, evidence = source_hash_score(); rows.append(result("R007", score, evidence))
        rows.append(result("R008", 1.0 if "descriptive cohort comparison" in text and not re.search(r"TODO|placeholder|nan", text, re.I) else 0.0, "scope note and no placeholders"))
    elif mode == "all_emails":
        expected = sorted((parse_email(path) for path in source_dir.glob("*.eml")), key=lambda item: item["sent_at"])
        table_rows = []
        if document and len(document.tables) == 1:
            table_rows = [[cell.text.strip() for cell in row.cells] for row in document.tables[0].rows[1:]]
        target_rows = [[item["sent_at"], item["sender"], item["subject"], item["body"]] for item in expected]
        matches = sum(1 for row in target_rows if row in table_rows)
        rows.append(result("R002", matches / len(target_rows), f"exact records matched={matches}/{len(target_rows)}"))
        actual_dates = [row[0] for row in table_rows]
        target_dates = [item["sent_at"] for item in expected]
        rows.append(result("R003", 1.0 if actual_dates == target_dates else 0.0, f"actual order={actual_dates}"))
        rows.append(result("R004", 1.0 if len(table_rows) == len(target_rows) and matches == len(target_rows) else 0.0, f"rows={len(table_rows)}/{len(target_rows)}"))
        rows.append(result("R005", 1.0 if "do not contain Date headers" in text and "Sent at" in text else 0.0, "explicit date-basis note"))
        rows.append(result("R006", 1.0 if document and len(document.tables) == 1 and len(document.tables[0].columns) == 4 else 0.0, "native four-column table"))
        score, evidence = source_hash_score(); rows.append(result("R007", score, evidence))
        rows.append(result("R008", 1.0 if native and not re.search(r"TODO|placeholder|nan", text, re.I) and all(all(cell for cell in row) for row in table_rows) else 0.0, "no blank fields or placeholders"))
    elif mode == "paper_summary":
        groups = {
            "R002": EXPECT["section_headings"],
            "R003": ["based solely on attention", "recurrence", "convolution", "encoder-decoder"],
            "R004": ["N = 6", "d_model = 512", "d_ff = 2048", "h = 8", "d_k = d_v = 64", "decoder mask", "sinusoidal positional encoding"],
            "R005": ["O(n^2 d)", "O(1) sequential operations", "O(1) maximum path length", "n is smaller than representation dimension d"],
            "R006": ["4.5 million", "36 million", "8 NVIDIA P100 GPUs", "100,000 steps", "12 hours", "300,000 steps", "3.5 days", "28.4 BLEU", "41.0 BLEU"],
            "R007": ["two machine-translation benchmarks", "very long sequences", "local or restricted attention", "images, audio, and video", "making generation less sequential"],
        }
        for identifier, anchors in groups.items():
            ratio, missing = anchor_ratio(text, anchors)
            rows.append(result(identifier, ratio, f"anchors={len(anchors)-len(missing)}/{len(anchors)}; missing={missing}"))
        source_text = subprocess.check_output(["pdftotext", str(source_dir / "paper.pdf"), "-"], text=True, errors="replace")
        source_anchors = ["Attention Is All You Need", "based solely on attention mechanisms", "8 NVIDIA P100 GPUs", "28.4 BLEU", "41.0"]
        source_ratio, source_missing = anchor_ratio(source_text, source_anchors)
        hash_score, hash_evidence = source_hash_score()
        rows.append(result("R008", min(source_ratio, hash_score), f"{hash_evidence}; source missing={source_missing}"))
        paragraph_count = len([p for p in document.paragraphs if compact(p.text)]) if document else 0
        rows.append(result("R009", min(1.0, paragraph_count / 12, (len(document.tables) / 2 if document else 0), len(compact(text)) / 2200), f"paragraphs={paragraph_count}; tables={len(document.tables) if document else 0}; chars={len(compact(text))}"))
        forbidden = [phrase for phrase in EXPECT["forbidden"] if compact(phrase) in compact(text)]
        rows.append(result("R010", 0.0 if forbidden else 1.0, f"forbidden={forbidden}", triggered=bool(forbidden)))
    return rows


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / EXPECT["output"]
    native = False
    document = None
    text = ""
    if output.is_file() and zipfile.is_zipfile(output):
        try:
            with zipfile.ZipFile(output) as archive:
                native = "word/document.xml" in archive.namelist()
            if native:
                document = Document(output)
                text, _, _ = document_content(document)
                native = bool(compact(text))
        except Exception:
            native = False
    results = [result("R001", 1.0 if native else 0.0, f"native non-empty DOCX={native}")]
    try:
        results.extend(evaluate_mode(document, text, native))
    except Exception as error:
        existing = {row["criterion_id"] for row in results}
        for spec in RUBRIC["criteria"]:
            if spec["id"] not in existing:
                results.append(result(spec["id"], 0.0, f"evaluation error: {error}"))
    by_id = {row["criterion_id"]: row for row in results}
    positive_total = sum(float(spec["weight"]) for spec in RUBRIC["criteria"] if float(spec["weight"]) > 0)
    earned = 0.0
    penalty = 0.0
    cap = 1.0
    hurdles = []
    for spec in RUBRIC["criteria"]:
        row = by_id.get(spec["id"], result(spec["id"], 0.0, "not evaluated"))
        weight = float(spec["weight"])
        if weight > 0:
            earned += weight * row["score"]
            if spec["type"] == "hurdle" and row["score"] + 1e-9 < float(spec.get("method_params", {}).get("minimum_score", 1.0)):
                hurdles.append(spec["id"])
                cap = min(cap, float(spec.get("score_cap_on_failure", 0.65)))
        elif row.get("raw", {}).get("triggered"):
            penalty += weight
            cap = min(cap, float(spec.get("method_params", {}).get("score_cap_on_trigger", 0.65)))
    reward = round(max(0.0, min(cap, (earned + penalty) / positive_total)), 6)
    return {
        "task_id": RUBRIC["task_id"],
        "rubric_version": RUBRIC["rubric_version"],
        "judge_version": RUBRIC["task_id"].split("/")[-1] + "-input-audited-v0.2",
        "artifact_sha256": [digest(output)] if output.is_file() else [],
        "criterion_results": sorted(results, key=lambda row: row["criterion_id"]),
        "reward": reward,
        "passed": reward >= float(RUBRIC["pass_threshold"]) and not hurdles,
        "hurdle_failures": hurdles,
    }


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    payload = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{payload['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)
