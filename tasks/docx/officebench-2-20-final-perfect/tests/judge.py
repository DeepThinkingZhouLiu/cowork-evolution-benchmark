from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from statistics import mean, median

from docx import Document
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


TASK_ID = "cowork-evo/officebench-2-20-final-perfect"
OUTPUT_NAME = "perfect_final_verification_packet.docx"
INPUT_SHA256 = "c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7"
RUBRIC_VERSION = "officebench-2-20-final-perfect-v0.3"
JUDGE_VERSION = "officebench-2-20-final-perfect-deterministic-v0.3"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
WEIGHTS = {"R001": .10, "R002": .20, "R003": .15, "R004": .15, "R005": .14, "R006": .12, "R007": .08, "R008": .06}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W_NS, "m": M_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def item(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def load_truth(path: Path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [clean(value) for value in next(ws.iter_rows(values_only=True))]
    rows = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        rows.append(record)
    wb.close()
    selected = [row for row in rows if isinstance(row["Final"], (int, float)) and float(row["Final"]) == 100]
    certificates = []
    for row in selected:
        scores = {name: int(row[name]) for name in ASSESSMENTS}
        pre = mean(scores[name] for name in ASSESSMENTS[:-1])
        overall = mean(scores.values())
        low, high = min(scores.values()), max(scores.values())
        rank = 1 + sum(int(other["Final"]) > scores["Final"] for other in rows)
        certificates.append({
            **row,
            "name": f"{row['First Name']} {row['Last Name']}",
            "scores": scores,
            "pre": pre,
            "overall": overall,
            "lowest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == low)} ({low})",
            "highest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == high)} ({high})",
            "range": high - low,
            "rank": f"{rank} of {len(rows)} (tie)",
            "lift": scores["Final"] - pre,
            "columns": {name: get_column_letter(headers.index(name) + 1) for name in ASSESSMENTS},
        })
    return rows, certificates


def matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, headers: list[str]) -> list[list[list[str]]]:
    return [data for table in document.tables if (data := matrix(table)) and data[0] == headers]


def ooxml_audit(path: Path):
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "DOCX 不存在或不是 ZIP/OOXML", {}
    try:
        with zipfile.ZipFile(path) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        prohibited = [name for name in entries if "embeddings/" in name.lower() or name.lower().endswith((".bin", ".exe"))]
        external = []
        for name, data in entries.items():
            if name.endswith(".rels"):
                root = etree.fromstring(data)
                external.extend(node.get("Target") for node in root.xpath("//*[local-name()='Relationship'][@TargetMode='External']"))
        ok = not prohibited and not external
        return ok, f"OOXML部件={len(entries)}；禁止嵌入={prohibited}；外部关系={external}", entries
    except Exception as exc:
        return False, f"OOXML解析失败: {exc}", {}


def render(path: Path):
    try:
        with tempfile.TemporaryDirectory(prefix="judge-final100-") as temp_dir:
            temp = Path(temp_dir)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            pdf = temp / f"{path.stem}.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            pages = int(re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE).group(1))
            width, height = map(float, re.search(r"^Page size:\s+([\d.]+) x ([\d.]+) pts", info, re.MULTILINE).groups())
            page_text = subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30).split("\f")
            page_text = [page for page in page_text if clean(page)]
            ok = pages == 4 and width < height and len(page_text) == 4
            return ok, f"LibreOffice pages={pages}, portrait={width < height}, nonblank={len(page_text)}"
    except Exception as exc:
        return False, f"LibreOffice渲染失败: {exc}"


def check_ledger(document: Document, rows: list[dict[str, object]]):
    tables = find_tables(document, ["Source row", "Student ID", "Name", "Final", "Final cell", "Decision", "Reason"])
    if len(tables) != 1:
        return 0.0, f"Population Decision Ledger 数量={len(tables)}"
    expected = []
    for row in rows:
        final = int(row["Final"])
        expected.append([f"Sheet1!{row['_row']}", clean(row["Student ID"]), f"{row['First Name']} {row['Last Name']}", str(final), f"Sheet1!I{row['_row']}", "Selected" if final == 100 else "Excluded", "Final = 100" if final == 100 else f"Final below 100 by {100-final}"])
    actual = tables[0][1:]
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = fraction(correct, len(expected) * 7)
    if len(actual) != len(expected):
        score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"Ledger cells={correct}/{len(expected)*7}; rows={len(actual)}/{len(expected)}; exact={actual == expected}"


def check_certificates(document: Document, certificates: list[dict[str, object]]):
    headings = [clean(p.text) for p in document.paragraphs if p.style and p.style.name.startswith("Heading") and clean(p.text).startswith("Verification Certificate")]
    expected_headings = [f"Verification Certificate • {cert['Student ID']}" for cert in certificates]
    identities = find_tables(document, ["Field", "Value", "Source cell"])
    assessments = find_tables(document, ["Assessment", "Score", "Source cell"])
    calculations = find_tables(document, ["Check", "Result", "Basis"])
    selection_checks = [headings == expected_headings, len(identities) == len(certificates), len(assessments) == len(certificates), len(calculations) == len(certificates)]
    selection_score = fraction(sum(selection_checks), len(selection_checks))

    evidence_correct = 0
    evidence_total = len(certificates) * 9
    calculation_correct = 0
    calculation_total = len(certificates) * 8
    for index, cert in enumerate(certificates):
        if index < len(identities):
            actual = {row[0]: row[1:] for row in identities[index][1:]}
            expected = {
                "Student ID": [clean(cert["Student ID"]), f"Sheet1!A{cert['_row']}"],
                "Name": [clean(cert["name"]), f"Sheet1!C{cert['_row']} + B{cert['_row']}"],
                "Academic Standing": [clean(cert["Academic Standing"]), f"Sheet1!D{cert['_row']}"],
                "Major": [clean(cert["Major"]), f"Sheet1!J{cert['_row']}"],
            }
            evidence_correct += sum(actual.get(key) == value for key, value in expected.items())
        if index < len(assessments):
            actual = {row[0]: row[1:] for row in assessments[index][1:]}
            expected = {name: [str(cert["scores"][name]), f"Sheet1!{cert['columns'][name]}{cert['_row']}"] for name in ASSESSMENTS}
            evidence_correct += sum(actual.get(key) == value for key, value in expected.items())
        if index < len(calculations):
            actual = {row[0]: row[1] for row in calculations[index][1:]}
            expected = {
                "Pre-Final average": f"{cert['pre']:.2f}",
                "Five-assessment average": f"{cert['overall']:.2f}",
                "Lowest assessment": clean(cert["lowest"]),
                "Highest assessment": clean(cert["highest"]),
                "Score range": str(cert["range"]),
                "Final rank": clean(cert["rank"]),
                "Exact predicate": "TRUE: Final = 100",
                "Final lift": f"+{cert['lift']:.2f}",
            }
            calculation_correct += sum(actual.get(key) == value for key, value in expected.items())
    return selection_score, f"Certificate headings={headings}; expected={expected_headings}; table counts={len(identities)}/{len(assessments)}/{len(calculations)}", fraction(evidence_correct, evidence_total), f"Certificate evidence={evidence_correct}/{evidence_total}", fraction(calculation_correct, calculation_total), f"Certificate calculations={calculation_correct}/{calculation_total}"


def check_equations(entries: dict[str, bytes], certificates: list[dict[str, object]]):
    root = etree.fromstring(entries["word/document.xml"])
    equations = [clean("".join(node.xpath(".//m:t/text()", namespaces=NS))) for node in root.xpath("//m:oMath", namespaces=NS)]
    expected = [f"100 - {cert['pre']:.2f} = {cert['lift']:.2f}" for cert in certificates]
    canonical = lambda value: clean(value).replace("−", "-").replace(" ", "")
    return [canonical(value) for value in equations] == [canonical(value) for value in expected], f"Native equations={equations}; expected={expected}"


def check_controls(document: Document, rows: list[dict[str, object]]):
    finals = [int(row["Final"]) for row in rows]
    perfect = [value for value in finals if value == 100]
    below = [value for value in finals if value < 100]
    control_tables = find_tables(document, ["Control", "Result", "Control", "Result"])
    expected_controls = [
        ["Source rows", str(len(rows)), "Numeric Final", str(len(finals))],
        ["Exact 100", str(len(perfect)), "Below 100", str(len(below))],
        ["Above 100", str(sum(value > 100 for value in finals)), "Missing / non-numeric", "0"],
        ["Perfect rate", f"{len(perfect)/len(rows)*100:.1f}%", "Final mean", f"{mean(finals):.1f}"],
        ["Final median", f"{median(finals):.1f}", "Next-highest", str(max(below))],
    ]
    distribution_tables = find_tables(document, ["Final band", "Count", "Share"])
    bands = [("0-19", 0, 19), ("20-39", 20, 39), ("40-59", 40, 59), ("60-79", 60, 79), ("80-99", 80, 99), ("100", 100, 100)]
    expected_distribution = []
    for label, lower, upper in bands:
        count = sum(lower <= value <= upper for value in finals)
        expected_distribution.append([label, str(count), f"{count/len(finals)*100:.1f}%"])
    control_ok = len(control_tables) == 1 and control_tables[0][1:] == expected_controls
    distribution_ok = len(distribution_tables) == 1 and distribution_tables[0][1:] == expected_distribution
    return fraction(int(control_ok) + int(distribution_ok), 2), f"Cohort Control={control_ok}; Final Distribution={distribution_ok}"


def check_boundary(document: Document, rows: list[dict[str, object]]):
    tables = find_tables(document, ["Source row", "Student ID", "Name", "Final", "Gap", "Final cell", "Decision"])
    if len(tables) != 1:
        return 0.0, f"Boundary table count={len(tables)}"
    excluded = sorted((row for row in rows if int(row["Final"]) < 100), key=lambda row: (-int(row["Final"]), clean(row["Student ID"])))[:3]
    expected = [[f"Sheet1!{row['_row']}", clean(row["Student ID"]), f"{row['First Name']} {row['Last Name']}", str(int(row["Final"])), str(100-int(row["Final"])), f"Sheet1!I{row['_row']}", "Excluded"] for row in excluded]
    actual = tables[0][1:]
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    return fraction(correct, 21) if len(actual) == 3 else fraction(correct, 21) * .75, f"Boundary cells={correct}/21; exact={actual == expected}"


def check_structure(document: Document, entries: dict[str, bytes], certificates: list[dict[str, object]], rows: list[dict[str, object]], render_ok: bool, render_evidence: str):
    root = etree.fromstring(entries["word/document.xml"])
    bookmarks = set(root.xpath("//w:bookmarkStart/@w:name", namespaces=NS))
    links = [(node.get(f"{{{W_NS}}}anchor"), clean("".join(node.xpath(".//w:t/text()", namespaces=NS)))) for node in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS)]
    expected_bookmarks = {f"certificate_{cert['Student ID']}" for cert in certificates}
    expected_links = {(bookmark, "Selected") for bookmark in expected_bookmarks}
    closure_tables = find_tables(document, ["Control", "Result", "Evidence"])
    selected_cells = ", ".join(f"Sheet1!I{cert['_row']}" for cert in certificates)
    expected_closure = [
        ["Population", f"{len(rows)} = {len(certificates)} Selected + {len(rows)-len(certificates)} Excluded", "Ledger reconciles"],
        ["Certificates", f"{len(certificates)} = {len(certificates)} Selected", "One certificate per selected row"],
        ["Selected Final cells", selected_cells, "Both cells equal 100"],
        ["Native equations", str(len(certificates)), "One Final-lift equation per certificate"],
        ["Above threshold", "0", "No Final above 100"],
        ["Missing / non-numeric", "0", "All nine Final values numeric"],
    ]
    closure_ok = len(closure_tables) == 1 and closure_tables[0][1:] == expected_closure
    headers = any(b"VERIFICATION PACKET" in data for name, data in entries.items() if name.startswith("word/header"))
    page_field = any(b"PAGE" in data for name, data in entries.items() if name.startswith("word/footer"))
    native_tables = len(document.tables) == 12 and len(root.xpath("//w:tblHeader", namespaces=NS)) >= 12 and len(root.xpath("//w:cantSplit", namespaces=NS)) >= 45
    no_media = not [name for name in entries if name.startswith("word/media/")]
    checks = [bookmarks == expected_bookmarks, set(links) == expected_links and len(links) == 2, closure_ok, headers and page_field, native_tables and no_media, render_ok]
    return fraction(sum(checks), len(checks)), f"Bookmarks={sorted(bookmarks)}; links={links}; Closure={closure_ok}; header/page={headers and page_field}; native/no-media={native_tables and no_media}; {render_evidence}"


def evaluate(artifact_dir: Path):
    source = Path(__file__).parents[1] / "data/input_files" / "CS161_Class_Roster.xlsx"
    output = artifact_dir / OUTPUT_NAME
    zip_ok, zip_evidence, entries = ooxml_audit(output)
    if not zip_ok:
        return [item(cid, 0, zip_evidence if cid == "R001" else "核心DOCX不可安全解析") for cid in WEIGHTS], output
    try:
        document = Document(output)
        rows, certificates = load_truth(source)
        render_ok, render_evidence = render(output)
        only_expected = sorted(path.name for path in artifact_dir.glob("*.docx")) == [OUTPUT_NAME]
        input_ok = hashlib.sha256(source.read_bytes()).hexdigest() == INPUT_SHA256
        native = len(document.tables) == 12 and len(clean(" ".join(cell.text for table in document.tables for row in table.rows for cell in row.cells))) > 700
        r1 = input_ok and only_expected and native and render_ok
        results = [item("R001", 1 if r1 else 0, f"{zip_evidence}; input={input_ok}; only_expected={only_expected}; native={native}; {render_evidence}")]
        score, evidence = check_ledger(document, rows)
        results.append(item("R002", score, evidence))
        r3, e3, r4, e4, r5_tables, e5_tables = check_certificates(document, certificates)
        results.append(item("R003", r3, e3))
        results.append(item("R004", r4, e4))
        equations_ok, equation_evidence = check_equations(entries, certificates)
        results.append(item("R005", (r5_tables + int(equations_ok)) / 2, f"{e5_tables}; {equation_evidence}"))
        r6, e6 = check_controls(document, rows)
        results.append(item("R006", r6, e6))
        r7, e7 = check_boundary(document, rows)
        results.append(item("R007", r7, e7))
        r8, e8 = check_structure(document, entries, certificates, rows, render_ok, render_evidence)
        results.append(item("R008", r8, e8))
        return results, output
    except Exception as exc:
        return [item(cid, 0, f"解析异常: {type(exc).__name__}: {exc}" if cid == "R001" else "解析异常，未计分") for cid in WEIGHTS], output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    results, output = evaluate(args.artifact_dir)
    values = {entry["criterion_id"]: float(entry["score"]) for entry in results}
    reward = sum(WEIGHTS[cid] * values.get(cid, 0) for cid in WEIGHTS)
    for cid, cap in (("R001", 0.0), ("R002", .35), ("R003", .45), ("R004", .65), ("R005", .76), ("R006", .78), ("R007", .78), ("R008", .78)):
        if values.get(cid, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    hurdles = ("R001", "R002", "R003", "R004", "R005", "R006", "R007", "R008")
    passed = reward >= .82 and all(values.get(cid, 0) == 1 for cid in hurdles)
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
