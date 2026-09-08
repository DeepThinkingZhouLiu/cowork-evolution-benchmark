from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-2-21-bill-courses"
OUTPUT_NAME = "bill_course_history_audit.docx"
INPUT_SHA256 = "d7764b78e2931fbcda87f2a45ef140693c3d6977e0235b264aa868cf949bab8a"
RUBRIC_VERSION = "officebench-2-21-bill-courses-v0.3"
JUDGE_VERSION = "officebench-2-21-bill-courses-deterministic-v0.3"
GRADE_POINTS = {"A": 4.0, "A-": 3.7, "B+": 3.3, "B": 3.0}
WEIGHTS = {"R001": .10, "R002": .19, "R003": .13, "R004": .17, "R005": .13, "R006": .12, "R007": .08, "R008": .08}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def item(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def load_truth(path: Path) -> list[dict[str, object]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [clean(value) for value in next(ws.iter_rows(values_only=True))]
    rows = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        record["_term"] = f"{record['Course Term']} {int(record['Year'])}"
        record["_points"] = GRADE_POINTS[clean(record["Course Grade"])]
        subject, number = clean(record["Course"]).split()
        record["_subject"] = subject
        record["_number"] = int(number)
        rows.append(record)
    wb.close()
    return rows


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, header: list[str]) -> list[list[list[str]]]:
    found = []
    for table in document.tables:
        matrix = table_matrix(table)
        if matrix and matrix[0] == header:
            found.append(matrix)
    return found


def exact_table(document: Document, header: list[str], expected: list[list[str]]) -> tuple[float, str]:
    tables = find_tables(document, header)
    if len(tables) != 1:
        return 0.0, f"table {header} count={len(tables)}"
    actual = tables[0][1:]
    total = sum(len(row) for row in expected)
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = fraction(correct, total)
    if len(actual) != len(expected):
        score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"rows={len(actual)}/{len(expected)}; cells={correct}/{total}; exact={actual == expected}"


def ooxml_audit(path: Path):
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "DOCX missing or not valid OOXML", {}
    try:
        with zipfile.ZipFile(path) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        prohibited = [name for name in entries if "embeddings/" in name.lower() or name.lower().endswith((".bin", ".exe", ".xlsm", ".docm"))]
        external = []
        for name, payload in entries.items():
            if name.endswith(".rels"):
                root = etree.fromstring(payload)
                external.extend(root.xpath("//*[local-name()='Relationship'][@TargetMode='External']/@Target"))
        ok = not prohibited and not external and "word/document.xml" in entries
        return ok, f"parts={len(entries)}; prohibited={prohibited}; external={external}", entries
    except Exception as exc:
        return False, f"OOXML parse failed: {exc}", {}


def render(path: Path):
    try:
        with tempfile.TemporaryDirectory(prefix="judge-bill-audit-") as temp_dir:
            temp = Path(temp_dir)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            pdf = temp / f"{path.stem}.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            pages = int(re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE).group(1))
            width, height = map(float, re.search(r"^Page size:\s+([\d.]+) x ([\d.]+) pts", info, re.MULTILINE).groups())
            page_text = subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30).split("\f")
            nonblank = sum(bool(clean(page)) for page in page_text)
            ok = pages == 4 and width < height and nonblank == 4
            return ok, f"LibreOffice pages={pages}; portrait={width < height}; nonblank={nonblank}"
    except Exception as exc:
        return False, f"LibreOffice render failed: {exc}"


def expected_term_rows(rows: list[dict[str, object]]) -> list[list[str]]:
    order = []
    grouped = defaultdict(list)
    for row in rows:
        term = clean(row["_term"])
        if term not in grouped:
            order.append(term)
        grouped[term].append(row)
    output = []
    cumulative_count = 0
    cumulative_points = 0.0
    for term in order:
        group = grouped[term]
        points = sum(float(row["_points"]) for row in group)
        cumulative_count += len(group)
        cumulative_points += points
        output.append([term, " · ".join(clean(row["Course"]) for row in group), str(len(group)), f"{points:.2f}", f"{points/len(group):.2f}", str(cumulative_count), f"{cumulative_points:.2f}", f"{cumulative_points/cumulative_count:.2f}"])
    return output


def expected_subject_rows(rows: list[dict[str, object]]) -> list[list[str]]:
    output = []
    for subject in ("CS", "MATH", "PHYS"):
        group = [row for row in rows if row["_subject"] == subject]
        points = sum(float(row["_points"]) for row in group)
        output.append([subject, str(len(group)), " → ".join(clean(row["Course"]) for row in group), f"{points:.2f}", f"{points/len(group):.2f}"])
    return output


def expected_continuity_rows(rows: list[dict[str, object]]) -> list[list[str]]:
    output = []
    for subject in ("CS", "MATH", "PHYS"):
        group = sorted((row for row in rows if row["_subject"] == subject), key=lambda row: int(row["_number"]))
        prior = None
        for row in group:
            number = int(row["_number"])
            if prior is None:
                output.append([subject, clean(row["Course"]), str(number), "—", "—", "START", f"Sheet1!C{row['_row']}"])
            else:
                delta = number - prior
                output.append([subject, clean(row["Course"]), str(number), str(prior), f"+{delta}", "CONTIGUOUS +1" if delta == 1 else f"JUMP +{delta}", f"Sheet1!C{row['_row']}"])
            prior = number
    return output


def check_ledger(document: Document, rows: list[dict[str, object]]):
    header = ["Source row", "Term", "Year", "Course", "Grade", "Points", "Source cells"]
    expected = [[f"Sheet1!{row['_row']}", clean(row["Course Term"]), str(int(row["Year"])), clean(row["Course"]), clean(row["Course Grade"]), f"{float(row['_points']):.1f}", f"Sheet1!A{row['_row']}:D{row['_row']}"] for row in rows]
    tables = find_tables(document, header)
    if len(tables) != 1:
        return 0.0, "ledger missing or duplicated", 0.0, "ledger missing or duplicated"
    actual = tables[0][1:]
    primary_total = len(expected) * 5
    primary_correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row[:5], target[:5]))
    primary_score = fraction(primary_correct, primary_total)
    if len(actual) != len(expected):
        primary_score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    evidence_total = len(expected) * 2
    evidence_correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row[5:7], target[5:7]))
    evidence_score = fraction(evidence_correct, evidence_total)
    if len(actual) != len(expected):
        evidence_score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return primary_score, f"chronology/data cells={primary_correct}/{primary_total}; rows={len(actual)}/{len(expected)}", evidence_score, f"point/locator cells={evidence_correct}/{evidence_total}"


def check_grade_policy(document: Document, ledger_score: float, ledger_evidence: str):
    scale = find_tables(document, ["Recorded grade", "A", "A-", "B+", "B"])
    scale_ok = len(scale) == 1 and scale[0][1:] == [["Grade points", "4.0", "3.7", "3.3", "3.0"]]
    return (ledger_score * 30 + int(scale_ok) * 5) / 35, f"{ledger_evidence}; grade-scale={scale_ok}"


def check_term(document: Document, rows: list[dict[str, object]]):
    return exact_table(document, ["Term", "Courses", "N", "Points", "Term GPA", "Cum. N", "Cum. points", "Cum. GPA"], expected_term_rows(rows))


def check_subject(document: Document, rows: list[dict[str, object]]):
    return exact_table(document, ["Subject", "N", "Source-order course sequence", "Points", "GPA"], expected_subject_rows(rows))


def check_continuity(document: Document, rows: list[dict[str, object]]):
    return exact_table(document, ["Subject", "Course", "No.", "Prior no.", "Delta", "Classification", "Course cell"], expected_continuity_rows(rows))


def check_closure(document: Document, rows: list[dict[str, object]]):
    counts = Counter(clean(row["Course Grade"]) for row in rows)
    distribution = [[grade, f"{GRADE_POINTS[grade]:.1f}", str(counts[grade]), f"{counts[grade]/len(rows)*100:.1f}%", f"{GRADE_POINTS[grade]*counts[grade]:.2f}"] for grade in ("A", "A-", "B+", "B")]
    dist_score, dist_evidence = exact_table(document, ["Grade", "Points", "Count", "Share", "Point contribution"], distribution)
    closure = [
        ["Source population", "15", "Rows 2–16; no blank records"],
        ["Term balance", "15 = 3 × 5", "Three records in each of five terms"],
        ["Subject balance", "15 = 5 + 5 + 5", "CS + MATH + PHYS"],
        ["Grade balance", "15 = 5 + 4 + 4 + 2", "A + A- + B+ + B"],
        ["Point balance", "54.00 = 19.00 + 17.60 + 17.40", "CS + MATH + PHYS"],
        ["Overall GPA", "3.60", "54.00 ÷ 15; unweighted"],
        ["Course-code edges", "12 = 11 contiguous + 1 jump", "Only MATH 152 → 250 is JUMP +98"],
        ["Exact duplicate rows", "0", "Term, Year, Course, Grade tuple"],
        ["Blank required cells", "0", "A2:D16"],
        ["Review controls", "5", "One unlocked rich-text control per term"],
    ]
    closure_score, closure_evidence = exact_table(document, ["Control", "Result", "Evidence / rule"], closure)
    return (dist_score + closure_score) / 2, f"distribution[{dist_evidence}]; closure[{closure_evidence}]"


def check_controls_and_structure(document: Document, entries: dict[str, bytes], render_ok: bool, render_evidence: str):
    root = etree.fromstring(entries["word/document.xml"])
    controls = root.xpath("//w:sdt", namespaces=NS)
    expected_tags = ["TermReview_Winter_2022", "TermReview_Spring_2022", "TermReview_Fall_2022", "TermReview_Winter_2023", "TermReview_Spring_2023"]
    actual_tags = [clean("".join(node.xpath("./w:sdtPr/w:tag/@w:val", namespaces=NS))) for node in controls]
    actual_text = [clean("".join(node.xpath("./w:sdtContent//w:t/text()", namespaces=NS))) for node in controls]
    unlocked = all(not node.xpath("./w:sdtPr/w:lock", namespaces=NS) for node in controls)
    rich_text = all(bool(node.xpath("./w:sdtPr/w:richText", namespaces=NS)) for node in controls)
    controls_ok = actual_tags == expected_tags and actual_text == ["Pending review"] * 5 and unlocked and rich_text
    headers = any(b"COURSE HISTORY AUDIT" in payload for name, payload in entries.items() if name.startswith("word/header"))
    page_field = any(b"PAGE" in payload for name, payload in entries.items() if name.startswith("word/footer"))
    native_tables = len(document.tables) == 9 and len(root.xpath("//w:tblHeader", namespaces=NS)) >= 9 and len(root.xpath("//w:cantSplit", namespaces=NS)) >= 68
    no_media = not [name for name in entries if name.startswith("word/media/")]
    headings = [clean(p.text) for p in document.paragraphs if p.style and p.style.name == "Title"]
    expected_headings = ["Bill Course History Audit Dossier", "Term Performance & Review Queue", "Subject Progression & Code Continuity", "Grade Distribution & Audit Closure"]
    body_text = clean(" ".join(paragraph.text for paragraph in document.paragraphs)).lower()
    prohibited_claims = ("completed all degree requirements", "approved to graduate", "credits earned", "prerequisite satisfied")
    no_invented_claim = not any(claim in body_text for claim in prohibited_claims)
    checks = [controls_ok, headers and page_field, native_tables and no_media, headings == expected_headings and no_invented_claim, render_ok]
    return fraction(sum(checks), len(checks)), f"controls={controls_ok}; tags={actual_tags}; unlocked={unlocked}; native/header/page/no-media={native_tables and headers and page_field and no_media}; titles={headings}; no-invented-claim={no_invented_claim}; {render_evidence}"


def evaluate(artifact_dir: Path):
    source = Path(__file__).parents[1] / "data/input_files" / "Bill_course_lists.xlsx"
    output = artifact_dir / OUTPUT_NAME
    safe, safe_evidence, entries = ooxml_audit(output)
    if not safe:
        return [item(cid, 0, safe_evidence if cid == "R001" else "core DOCX unavailable") for cid in WEIGHTS], output
    try:
        document = Document(output)
        rows = load_truth(source)
        render_ok, render_evidence = render(output)
        only_expected = sorted(path.name for path in artifact_dir.glob("*.docx")) == [OUTPUT_NAME]
        input_ok = hashlib.sha256(source.read_bytes()).hexdigest() == INPUT_SHA256
        native = len(document.tables) == 9 and sum(len(table.rows) for table in document.tables) >= 60
        r1 = input_ok and only_expected and native and render_ok
        results = [item("R001", 1 if r1 else 0, f"{safe_evidence}; input={input_ok}; only_expected={only_expected}; native={native}; {render_evidence}")]
        r2, e2, ledger_evidence_score, ledger_evidence = check_ledger(document, rows)
        results.append(item("R002", r2, e2))
        r3, e3 = check_grade_policy(document, ledger_evidence_score, ledger_evidence)
        results.append(item("R003", r3, e3))
        r4, e4 = check_term(document, rows)
        results.append(item("R004", r4, e4))
        r5, e5 = check_subject(document, rows)
        results.append(item("R005", r5, e5))
        r6, e6 = check_continuity(document, rows)
        results.append(item("R006", r6, e6))
        r7, e7 = check_closure(document, rows)
        results.append(item("R007", r7, e7))
        r8, e8 = check_controls_and_structure(document, entries, render_ok, render_evidence)
        results.append(item("R008", r8, e8))
        return results, output
    except Exception as exc:
        return [item(cid, 0, f"parse error: {type(exc).__name__}: {exc}" if cid == "R001" else "parse error; not scored") for cid in WEIGHTS], output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    results, output = evaluate(args.artifact_dir)
    values = {entry["criterion_id"]: float(entry["score"]) for entry in results}
    reward = sum(WEIGHTS[cid] * values.get(cid, 0) for cid in WEIGHTS)
    for cid, cap in (("R001", 0.0), ("R002", .35), ("R003", .52), ("R004", .65), ("R005", .72), ("R006", .76), ("R007", .78), ("R008", .80)):
        if values.get(cid, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    passed = reward >= .84 and all(values.get(cid, 0) == 1 for cid in WEIGHTS)
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
