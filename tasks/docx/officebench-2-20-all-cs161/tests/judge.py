#!/usr/bin/env python3
"""Deterministic Judge for the CS161 roster control book."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-2-20-all-cs161"
OUTPUT_NAME = "cs161_roster_control_book.docx"
RUBRIC_VERSION = "officebench-2-20-all-cs161-v0.3"
JUDGE_VERSION = "officebench-2-20-all-cs161-deterministic-v0.3"
SOURCE_SHA256 = "c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
SPECS = {
    "R001": (10, "hurdle", 0.0),
    "R002": (23, "hurdle", 0.35),
    "R003": (8, "positive", None),
    "R004": (12, "hurdle", 0.68),
    "R005": (13, "hurdle", 0.68),
    "R006": (10, "hurdle", 0.72),
    "R007": (9, "positive", None),
    "R008": (7, "positive", None),
    "R009": (4, "positive", None),
    "R010": (4, "positive", None),
}


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def item(identifier: str, fraction: float, evidence: str) -> dict:
    score = max(0.0, min(1.0, fraction))
    return {
        "criterion_id": identifier,
        "score": score,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"passed": score == 1.0, "weight": SPECS[identifier][0], "type": SPECS[identifier][1]},
    }


def source_root() -> Path:
    mounted = Path("/app/input_files")
    if (mounted / "CS161_Class_Roster.xlsx").is_file():
        return mounted
    return Path(__file__).resolve().parents[1] / "data/input_files"


def source_truth() -> tuple[list[str], list[dict]]:
    workbook = load_workbook(source_root() / "CS161_Class_Roster.xlsx", data_only=True, read_only=True)
    sheet = workbook["Sheet1"]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [norm(value) for value in rows[0]]
    records = []
    for source_row, values in enumerate(rows[1:], start=2):
        if any(value is not None for value in values):
            records.append({**dict(zip(headers, values)), "Source row": f"Sheet1!{source_row}"})
    return headers, records


def table_by_header(document, expected: list[str]):
    wanted = [norm(value).lower() for value in expected]
    for table in document.tables:
        if not table.rows:
            continue
        actual = [norm(cell.text).lower() for cell in table.rows[0].cells]
        if actual == wanted:
            return table
    return None


def table_values(table) -> list[list[str]]:
    if table is None:
        return []
    return [[norm(cell.text) for cell in row.cells] for row in table.rows[1:]]


def render_check(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    pdfinfo = shutil.which("pdfinfo")
    if not all((libreoffice, pdftotext, pdfinfo)):
        return False, "LibreOffice/Poppler dependencies are unavailable"
    with tempfile.TemporaryDirectory(prefix="cs161-roster-judge-") as temp:
        root = Path(temp)
        profile = (root / "profile").resolve().as_uri()
        converted = subprocess.run(
            [libreoffice, "--headless", f"-env:UserInstallation={profile}", "--convert-to", "pdf", "--outdir", temp, str(path)],
            capture_output=True,
            text=True,
            timeout=240,
            check=False,
        )
        pdf = root / f"{path.stem}.pdf"
        if converted.returncode or not pdf.is_file():
            return False, f"PDF conversion failed: {(converted.stderr or converted.stdout)[-400:]}"
        extracted = subprocess.run([pdftotext, "-layout", str(pdf), "-"], capture_output=True, text=True, timeout=120, check=False)
        info = subprocess.run([pdfinfo, str(pdf)], capture_output=True, text=True, timeout=60, check=False)
        if extracted.returncode or info.returncode:
            return False, "PDF text/page inspection failed"
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        detailed = subprocess.run([pdfinfo, "-f", "1", "-l", str(pages), str(pdf)], capture_output=True, text=True, timeout=60, check=False)
        page_blocks = extracted.stdout.split("\f")[:pages]
        counts = [len(re.sub(r"\s+", "", block)) for block in page_blocks]
        sizes = [(float(width), float(height)) for width, height in re.findall(r"^Page\s+\d+\s+size:\s+([\d.]+)\s+x\s+([\d.]+)", detailed.stdout, flags=re.MULTILINE)]
        orientation_ok = pages >= 4 and len(sizes) == pages and all(width < height for width, height in sizes[:2]) and sizes[2][0] > sizes[2][1] and all(width < height for width, height in sizes[3:])
        required = ["CS161 Roster Control Book", "Source-order Roster Register", "Assessment Reconciliation", "Alphabetical Name Index"]
        missing = [value for value in required if value not in extracted.stdout]
        blank_pages = [index + 1 for index, count in enumerate(counts) if count < 35]
        ok = 4 <= pages <= 8 and len(counts) == pages and not blank_pages and not missing and orientation_ok
        return ok, f"pages={pages}, page_characters={counts}, sizes={sizes}, orientation_ok={orientation_ok}, blank_pages={blank_pages}, missing={missing}"


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    results: list[dict] = []
    if not output.is_file() or not zipfile.is_zipfile(output):
        results.append(item("R001", 0, f"Missing or corrupt {OUTPUT_NAME}"))
        return finalize(results, output)

    headers, records = source_truth()
    digest = hashlib.sha256((source_root() / "CS161_Class_Roster.xlsx").read_bytes()).hexdigest()
    try:
        document = Document(output)
        with zipfile.ZipFile(output) as archive:
            names = archive.namelist()
            document_xml = archive.read("word/document.xml")
            rel_xmls = [archive.read(name) for name in names if name.endswith(".rels")]
    except Exception as error:
        results.append(item("R001", 0, f"DOCX parse failed: {error}"))
        return finalize(results, output)

    native_ok = digest == SOURCE_SHA256 and len(document.tables) >= 5 and sum(bool(norm(p.text)) for p in document.paragraphs) >= 12
    results.append(item("R001", 1 if native_ok else 0, f"source_sha={digest}; paragraphs={len(document.paragraphs)}; tables={len(document.tables)}"))

    roster_headers = ["Source row", "Student ID", "Last name", "First name", "Academic standing", "Major", *ASSESSMENTS]
    roster = table_by_header(document, roster_headers)
    expected_roster = [
        [
            record["Source row"], record["Student ID"], record["Last Name"], record["First Name"], record["Academic Standing"],
            record["Major"], *[record[name] for name in ASSESSMENTS],
        ]
        for record in records
    ]
    expected_roster = [[norm(value) for value in row] for row in expected_roster]
    actual_roster = table_values(roster)
    matched_rows = sum(actual == expected for actual, expected in zip(actual_roster, expected_roster))
    roster_exact = actual_roster == expected_roster
    roster_fraction = 1.0 if roster_exact else matched_rows / max(len(expected_roster), len(actual_roster), 1)
    results.append(item("R002", roster_fraction, f"exact_rows={matched_rows}/{len(expected_roster)}; actual_rows={len(actual_roster)}"))

    expected_ids = [norm(record["Student ID"]) for record in records]
    actual_ids = [row[1] for row in actual_roster if len(row) > 1]
    locators = [row[0] for row in actual_roster if row]
    expected_locators = [record["Source row"] for record in records]
    trace_parts = [actual_ids == expected_ids, locators == expected_locators, len(actual_ids) == len(set(actual_ids))]
    results.append(item("R003", sum(trace_parts) / len(trace_parts), f"ids_exact={trace_parts[0]}; locators_exact={trace_parts[1]}; unique={trace_parts[2]}"))

    distribution = table_by_header(document, ["Dimension", "Category", "Count", "Share"])
    expected_distribution = []
    total = len(records)
    for dimension, key in (("Major", "Major"), ("Academic standing", "Academic Standing")):
        counts = Counter(norm(record[key]) for record in records)
        expected_distribution.extend([[dimension, category, str(counts[category]), f"{counts[category] / total * 100:.1f}%"] for category in sorted(counts)])
    actual_distribution = table_values(distribution)
    summary_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    distribution_ok = actual_distribution == expected_distribution and f"source contains {total} populated student rows" in summary_text.lower()
    results.append(item("R004", 1 if distribution_ok else 0, f"distribution_rows={len(actual_distribution)}/{len(expected_distribution)}; total_statement={total}"))

    assessment = table_by_header(document, ["Assessment", "Valid N", "Average", "Minimum", "Maximum"])
    expected_assessment = []
    for name in ASSESSMENTS:
        values = [int(record[name]) for record in records]
        expected_assessment.append([name, str(len(values)), f"{sum(values) / len(values):.2f}", str(min(values)), str(max(values))])
    actual_assessment = table_values(assessment)
    assessment_ok = actual_assessment == expected_assessment
    results.append(item("R005", 1 if assessment_ok else 0, f"actual={actual_assessment}; expected={expected_assessment}"))

    index_table = table_by_header(document, ["Name", "Student ID", "Source row", "Roster link"])
    expected_index = [
        [f"{record['Last Name']}, {record['First Name']}", norm(record["Student ID"]), record["Source row"]]
        for record in sorted(records, key=lambda record: (norm(record["Last Name"]), norm(record["First Name"])))
    ]
    actual_index = [row[:3] for row in table_values(index_table)]
    root = etree.fromstring(document_xml)
    bookmarks = set(root.xpath("//w:bookmarkStart/@w:name", namespaces=W_NS))
    anchors = root.xpath("//w:hyperlink/@w:anchor", namespaces=W_NS)
    expected_anchors = [f"roster_{record['Student ID']}" for record in sorted(records, key=lambda record: (norm(record["Last Name"]), norm(record["First Name"])))]
    index_parts = [actual_index == expected_index, all(anchor in bookmarks for anchor in expected_anchors), Counter(anchors) == Counter(expected_anchors)]
    results.append(item("R006", sum(index_parts) / len(index_parts), f"content={index_parts[0]}; bookmarks={len(bookmarks & set(expected_anchors))}/9; exact_links={index_parts[2]}"))

    quality = table_by_header(document, ["Control", "Result", "Evidence"])
    expected_quality = [
        ["Complete source rows", "PASS", "9 of 9 rows have all 10 fields"],
        ["Unique Student IDs", "PASS", "9 unique of 9 IDs"],
        ["Student ID format", "PASS", "9 of 9 match S + 4 digits"],
        ["Integer score bounds", "PASS", "45 of 45 scores are integers in 0–100"],
    ]
    actual_quality = table_values(quality)
    quality_matches = sum(actual == expected for actual, expected in zip(actual_quality, expected_quality))
    results.append(item("R007", quality_matches / len(expected_quality), f"exact_controls={quality_matches}/{len(expected_quality)}"))

    toc_instruction = " ".join(root.xpath("//w:instrText/text() | //w:fldSimple/@w:instr", namespaces=W_NS))
    headings = {norm(p.text): p.style.name if p.style else "" for p in document.paragraphs if norm(p.text)}
    orientations = [section.orientation for section in document.sections]
    headers_text = [norm(" ".join(p.text for p in section.header.paragraphs)) for section in document.sections]
    control_sentence = f"Control total: {total} source rows = {total} roster rows = {total} index entries."
    structure_parts = [
        len(document.sections) == 3,
        orientations == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT],
        "TOC" in toc_instruction and "1-2" in toc_instruction,
        headings.get("CS161 Roster Control Book") == "Title",
        all(headings.get(value) == "Heading 1" for value in ("Contents", "Control Summary", "Source-order Roster Register", "Assessment Reconciliation", "Data-quality Controls", "Alphabetical Name Index")),
        all(value == "CS161 Roster Control Book" for value in headers_text),
        control_sentence in summary_text,
    ]
    results.append(item("R008", sum(structure_parts) / len(structure_parts), f"sections={len(document.sections)}; orientations={orientations}; TOC={toc_instruction!r}; parts={structure_parts}"))

    external_targets = []
    for payload in rel_xmls:
        rel_root = etree.fromstring(payload)
        external_targets.extend(rel_root.xpath("//r:Relationship[@TargetMode='External']/@Target", namespaces=REL_NS))
    header_rows = len(root.xpath("//w:tr[w:trPr/w:tblHeader]", namespaces=W_NS))
    nonsplit_rows = len(root.xpath("//w:tr[w:trPr/w:cantSplit]", namespaces=W_NS))
    unsafe = [name for name in names if name.lower().endswith("vbaproject.bin") or "embeddings/" in name.lower()]
    image_count = len([name for name in names if name.startswith("word/media/")])
    editable_parts = [len(document.tables) == 5, header_rows >= 5, nonsplit_rows >= 35, not unsafe, not external_targets, image_count == 0]
    results.append(item("R009", sum(editable_parts) / len(editable_parts), f"tables={len(document.tables)}; header_rows={header_rows}; non_split={nonsplit_rows}; unsafe={unsafe}; external={external_targets}; images={image_count}"))

    render_ok, render_evidence = render_check(output)
    results.append(item("R010", 1 if render_ok else 0, render_evidence))
    return finalize(results, output)


def finalize(results: list[dict], output: Path) -> dict:
    scores = {result["criterion_id"]: result["score"] for result in results}
    reward = sum(weight * scores.get(identifier, 0.0) for identifier, (weight, _, _) in SPECS.items()) / sum(spec[0] for spec in SPECS.values())
    hurdle_failures = []
    caps = []
    for identifier, (_, kind, cap) in SPECS.items():
        if kind == "hurdle" and scores.get(identifier, 0.0) != 1.0:
            hurdle_failures.append(identifier)
            if cap is not None:
                caps.append(cap)
    if caps:
        reward = min(reward, min(caps))
    reward = round(max(0.0, min(1.0, reward)), 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    payload = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{payload['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
