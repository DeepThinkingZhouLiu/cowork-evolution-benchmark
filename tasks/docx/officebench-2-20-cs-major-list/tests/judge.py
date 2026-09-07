from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from statistics import mean

from docx import Document
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


TASK_ID = "cowork-evo/officebench-2-20-cs-major-list"
OUTPUT_NAME = "cs_major_advising_docket.docx"
RUBRIC_VERSION = "officebench-2-20-cs-major-list-v0.3"
JUDGE_VERSION = "officebench-2-20-cs-major-list-deterministic-v0.3"
INPUT_SHA256 = "c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
WEIGHTS = {"R001": 0.10, "R002": 0.18, "R003": 0.16, "R004": 0.16, "R005": 0.14, "R006": 0.12, "R007": 0.08, "R008": 0.06}
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def result(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def ratio(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def tier(final_score: int) -> tuple[int, str, str]:
    if final_score <= 5:
        return 1, "Priority 1", "0-5"
    if final_score <= 20:
        return 2, "Priority 2", "6-20"
    return 3, "Priority 3", "21-100"


def source_truth(path: Path) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [clean(value) for value in next(ws.iter_rows(values_only=True))]
    rows: list[dict[str, object]] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        rows.append(record)
    wb.close()

    selected: list[dict[str, object]] = []
    for record in rows:
        if clean(record["Major"]) != "CS":
            continue
        scores = {name: int(record[name]) for name in ASSESSMENTS}
        values = list(scores.values())
        delta = scores["Final"] - scores["Midterm"]
        tier_number, tier_name, band = tier(scores["Final"])
        lowest_value, highest_value = min(values), max(values)
        selected.append({
            **record,
            "name": f"{record['First Name']} {record['Last Name']}",
            "scores": scores,
            "average": round(mean(values), 1),
            "lowest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == lowest_value)} ({lowest_value})",
            "highest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == highest_value)} ({highest_value})",
            "range": highest_value - lowest_value,
            "delta": delta,
            "trend": "Decline" if delta < 0 else "Increase" if delta > 0 else "Flat",
            "tier_number": tier_number,
            "tier": tier_name,
            "band": band,
            "assessment_columns": {name: get_column_letter(headers.index(name) + 1) for name in ASSESSMENTS},
        })
    queue = sorted(selected, key=lambda item: (int(item["tier_number"]), int(item["scores"]["Final"]), clean(item["Student ID"])))
    return rows, selected, queue


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def tables_with_header(document: Document, header: list[str]) -> list[list[list[str]]]:
    expected = [clean(value) for value in header]
    matches = []
    for table in document.tables:
        matrix = table_matrix(table)
        if matrix and matrix[0] == expected:
            matches.append(matrix)
    return matches


def zip_audit(path: Path) -> tuple[bool, str, dict[str, bytes]]:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "DOCX 不存在或不是 ZIP/OOXML", {}
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            entries = {name: archive.read(name) for name in names}
            bad = [name for name in names if name.lower().endswith((".bin", ".vba", ".vbaProject", ".exe")) or "embeddings/" in name.lower()]
            external = []
            for name, data in entries.items():
                if not name.endswith(".rels"):
                    continue
                root = etree.fromstring(data)
                for relationship in root.xpath("//*[local-name()='Relationship']"):
                    if relationship.get("TargetMode") == "External":
                        external.append(f"{name}:{relationship.get('Target')}")
            if bad or external:
                return False, f"禁止的嵌入/宏={bad}; 外部关系={external}", entries
            return True, f"OOXML ZIP 正常；{len(names)} 个部件；无宏、OLE、嵌入文件或外部关系", entries
    except Exception as exc:
        return False, f"OOXML 审计失败: {exc}", {}


def render_audit(path: Path) -> tuple[bool, str]:
    try:
        with tempfile.TemporaryDirectory(prefix="judge-cs-major-") as temp_dir:
            temp = Path(temp_dir)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=120)
            pdf = temp / f"{path.stem}.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            page_match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
            size_match = re.search(r"^Page size:\s+([\d.]+) x ([\d.]+) pts", info, re.MULTILINE)
            pages = int(page_match.group(1)) if page_match else 0
            portrait = bool(size_match and float(size_match.group(1)) < float(size_match.group(2)))
            text = subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30)
            page_texts = text.split("\f")
            while page_texts and not clean(page_texts[-1]):
                page_texts.pop()
            nonblank = len(page_texts) == pages and all(clean(page) for page in page_texts)
            ok = pages == 5 and portrait and nonblank
            return ok, f"LibreOffice PDF pages={pages}, portrait={portrait}, nonblank_pages={sum(bool(clean(page)) for page in page_texts)}"
    except Exception as exc:
        return False, f"LibreOffice 渲染失败: {exc}"


def criterion_queue_filter(document: Document, all_text: str, selected: list[dict[str, object]], queue: list[dict[str, object]], rows: list[dict[str, object]]) -> tuple[float, str, list[list[str]]]:
    tables = tables_with_header(document, ["Order", "Student ID", "Name", "Standing", "Final", "Δ M->F", "Tier", "Source row", "Evidence"])
    if len(tables) != 1:
        return 0.0, f"Review Queue 表数量={len(tables)}，应为1", []
    actual = tables[0][1:]
    expected_by_id = {
        clean(student["Student ID"]): [clean(student["Student ID"]), clean(student["name"]), clean(student["Academic Standing"]), str(student["scores"]["Final"]), f"Sheet1!{student['_row']}"]
        for student in selected
    }
    actual_by_id = {row[1]: [row[1], row[2], row[3], row[4], row[7]] for row in actual if len(row) == 9}
    matches = sum(actual_by_id.get(student_id) == values for student_id, values in expected_by_id.items())
    no_extra = set(actual_by_id) == set(expected_by_id) and len(actual) == len(expected_by_id)
    non_cs = [row for row in rows if clean(row["Major"]) != "CS"]
    leaked = []
    for row in non_cs:
        for term in (clean(row["Student ID"]), f"{row['First Name']} {row['Last Name']}", clean(row["First Name"]), clean(row["Last Name"])):
            if term and re.search(rf"(?<!\w){re.escape(term)}(?!\w)", all_text, re.IGNORECASE):
                leaked.append(term)
    score = ratio(matches, len(expected_by_id))
    if not no_extra or leaked:
        score = min(score, 0.75)
    return score, f"CS 身份字段匹配={matches}/{len(expected_by_id)}；集合/行数准确={no_extra}；非CS个体泄漏={sorted(set(leaked))}", actual


def criterion_queue_logic(actual: list[list[str]], queue: list[dict[str, object]]) -> tuple[float, str]:
    expected = [[str(order), clean(student["Student ID"]), clean(student["name"]), clean(student["Academic Standing"]), str(student["scores"]["Final"]), f"{student['delta']:+d}", clean(student["tier"]), f"Sheet1!{student['_row']}", "Open evidence record"] for order, student in enumerate(queue, start=1)]
    if not expected:
        return 0.0, "输入未产生预期队列"
    row_matches = sum(row == target for row, target in zip(actual, expected))
    cell_total = len(expected) * len(expected[0])
    cell_matches = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    exact = actual == expected
    score = ratio(cell_matches, cell_total)
    if len(actual) != len(expected):
        score *= ratio(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"完整队列行={row_matches}/{len(expected)}；单元格={cell_matches}/{cell_total}；顺序和逻辑完全准确={exact}"


def criterion_profiles(document: Document, queue: list[dict[str, object]]) -> tuple[float, str, float, str]:
    identities = tables_with_header(document, ["Field", "Value", "Source"])
    assessments = tables_with_header(document, ["Assessment", "Score", "Source cell"])
    derived = tables_with_header(document, ["Derived check", "Result", "Calculation basis"])
    evidence_correct = 0
    evidence_total = len(queue) * 9
    derived_correct = 0
    derived_total = len(queue) * 7
    for index, student in enumerate(queue):
        if index < len(identities):
            actual = {row[0]: row[1:] for row in identities[index][1:] if len(row) == 3}
            expected = {
                "Student ID": [clean(student["Student ID"]), f"Sheet1!A{student['_row']}"],
                "Name": [clean(student["name"]), f"Sheet1!C{student['_row']} + B{student['_row']}"],
                "Academic Standing": [clean(student["Academic Standing"]), f"Sheet1!D{student['_row']}"],
                "Major filter": ["CS", f"Sheet1!J{student['_row']}"],
            }
            evidence_correct += sum(actual.get(key) == value for key, value in expected.items())
        if index < len(assessments):
            actual = {row[0]: row[1:] for row in assessments[index][1:] if len(row) == 3}
            expected = {name: [str(student["scores"][name]), f"Sheet1!{student['assessment_columns'][name]}{student['_row']}"] for name in ASSESSMENTS}
            evidence_correct += sum(actual.get(key) == value for key, value in expected.items())
        if index < len(derived):
            actual = {row[0]: row[1] for row in derived[index][1:] if len(row) == 3}
            expected = {
                "Five-assessment average": f"{student['average']:.1f}",
                "Lowest assessment": clean(student["lowest"]),
                "Highest assessment": clean(student["highest"]),
                "Score range": str(student["range"]),
                "Midterm -> Final": f"{student['delta']:+d}",
                "Trend": clean(student["trend"]),
                "Review tier": clean(student["tier"]),
            }
            derived_correct += sum(actual.get(key) == value for key, value in expected.items())
    evidence_score = ratio(evidence_correct, evidence_total)
    derived_score = ratio(derived_correct, derived_total)
    return evidence_score, f"身份与评分证据字段={evidence_correct}/{evidence_total}；identity tables={len(identities)}，assessment tables={len(assessments)}", derived_score, f"逐学生派生检查={derived_correct}/{derived_total}；derived tables={len(derived)}"


def criterion_appendix(document: Document, rows: list[dict[str, object]], selected: list[dict[str, object]]) -> tuple[float, str]:
    tables = tables_with_header(document, ["Assessment", "CS N", "CS avg", "Non-CS N", "Non-CS avg", "CS - non-CS"])
    if len(tables) != 1:
        return 0.0, f"Cohort Comparison 表数量={len(tables)}，应为1"
    selected_ids = {student["Student ID"] for student in selected}
    cs_rows = [row for row in rows if row["Student ID"] in selected_ids]
    non_cs_rows = [row for row in rows if row["Student ID"] not in selected_ids]
    expected = []
    for name in ASSESSMENTS:
        cs_average = mean(float(row[name]) for row in cs_rows)
        non_cs_average = mean(float(row[name]) for row in non_cs_rows)
        expected.append([name, str(len(cs_rows)), f"{cs_average:.1f}", str(len(non_cs_rows)), f"{non_cs_average:.1f}", f"{cs_average - non_cs_average:+.1f}"])
    actual = tables[0][1:]
    cells = len(expected) * 6
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = ratio(correct, cells)
    if len(actual) != len(expected):
        score *= ratio(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"对照附录单元格={correct}/{cells}；行数={len(actual)}/{len(expected)}"


def criterion_native_objects(entries: dict[str, bytes], queue: list[dict[str, object]]) -> tuple[float, str]:
    try:
        document_root = etree.fromstring(entries["word/document.xml"])
        bookmarks = set(document_root.xpath("//w:bookmarkStart/@w:name", namespaces=NS))
        expected_bookmarks = {f"student_{student['Student ID']}" for student in queue}
        links = []
        for node in document_root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS):
            links.append((node.get(f"{{{NS['w']}}}anchor"), clean("".join(node.xpath(".//w:t/text()", namespaces=NS)))))
        expected_links = {(bookmark, "Open evidence record") for bookmark in expected_bookmarks}

        comments_root = etree.fromstring(entries["word/comments.xml"])
        comments = {}
        for node in comments_root.xpath("//w:comment", namespaces=NS):
            comment_id = node.get(f"{{{NS['w']}}}id")
            comments[comment_id] = clean("".join(node.xpath(".//w:t/text()", namespaces=NS)))
        expected_comments = {
            f"Tier rule: Final {student['band']} -> {student['tier']}. Observed Final: {student['scores']['Final']}." for student in queue
        }
        anchors = {}
        for start in document_root.xpath("//w:commentRangeStart", namespaces=NS):
            comment_id = start.get(f"{{{NS['w']}}}id")
            cells = start.xpath("ancestor::w:tc[1]", namespaces=NS)
            anchors[comment_id] = clean("".join(cells[0].xpath(".//w:t/text()", namespaces=NS))) if cells else ""
        comment_refs = set(document_root.xpath("//w:commentReference/@w:id", namespaces=NS))
        correct_comments = set(comments.values()) == expected_comments and len(comments) == len(queue)
        anchored = len(anchors) == len(queue) and all(anchors.get(comment_id) in {"Priority 1", "Priority 2", "Priority 3"} for comment_id in comments) and set(comments) == comment_refs

        checks = [bookmarks == expected_bookmarks, set(links) == expected_links and len(links) == len(queue), correct_comments, anchored]
        return ratio(sum(checks), len(checks)), f"书签={sorted(bookmarks)}；内部链接={links}；批注数={len(comments)}；批注文本准确={correct_comments}；Tier锚定与引用闭环={anchored}"
    except Exception as exc:
        return 0.0, f"原生书签/链接/批注解析失败: {exc}"


def criterion_structure(document: Document, entries: dict[str, bytes], render_ok: bool, render_evidence: str, rows: list[dict[str, object]], selected: list[dict[str, object]]) -> tuple[float, str]:
    controls = tables_with_header(document, ["Control", "Result", "Evidence"])
    selected_source_rows = ", ".join(f"Sheet1!{student['_row']}" for student in sorted(selected, key=lambda item: int(item["_row"])))
    expected_controls = {
        "Source row count": str(len(rows)),
        "CS selected count": str(len(selected)),
        "Non-CS excluded count": str(len(rows) - len(selected)),
        "Selected source rows": selected_source_rows,
        "Tier rules": "0-5 -> P1; 6-20 -> P2; 21-100 -> P3",
        "Ranking rule": "Tier -> Final asc -> Student ID asc",
        "Reconciliation": f"{len(selected)} queue rows = {len(selected)} evidence records",
    }
    control_ok = False
    if len(controls) == 1:
        actual = {row[0]: row[1] for row in controls[0][1:] if len(row) == 3}
        control_ok = actual == expected_controls

    heading_text = {clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style and paragraph.style.name.startswith("Heading")}
    required_headings = {"Review Queue", "Decision Rule Card", "Cohort Comparison Appendix", "Control Summary"} | {f"Evidence Record • {student['Student ID']}" for student in selected}
    heading_ok = required_headings.issubset(heading_text)
    document_root = etree.fromstring(entries.get("word/document.xml", b"<x/>"))
    repeat_headers = len(document_root.xpath("//w:tblHeader", namespaces=NS))
    nonsplit_rows = len(document_root.xpath("//w:cantSplit", namespaces=NS))
    footer_parts = [data for name, data in entries.items() if name.startswith("word/footer") and name.endswith(".xml")]
    header_parts = [data for name, data in entries.items() if name.startswith("word/header") and name.endswith(".xml")]
    page_field = any(b"PAGE" in data for data in footer_parts)
    header_ok = any(b"SOURCE-CONTROLLED" in data for data in header_parts)
    media = [name for name in entries if name.startswith("word/media/")]
    no_flattened = not media and len(document.tables) >= 13
    native_table_controls = repeat_headers >= 13 and nonsplit_rows >= 45
    checks = [control_ok, heading_ok, page_field and header_ok, native_table_controls, no_flattened, render_ok]
    return ratio(sum(checks), len(checks)), f"Control={control_ok}；Heading={heading_ok}；页眉/页码={header_ok and page_field}；重复表头={repeat_headers}；不拆行={nonsplit_rows}；无图片且原生表格={no_flattened}；{render_evidence}"


def evaluate(artifact_dir: Path) -> tuple[list[dict[str, object]], Path]:
    source = Path(__file__).parents[1] / "data/input_files" / "CS161_Class_Roster.xlsx"
    output = artifact_dir / OUTPUT_NAME
    input_digest = hashlib.sha256(source.read_bytes()).hexdigest()
    only_expected = sorted(path.name for path in artifact_dir.glob("*.docx")) == [OUTPUT_NAME]
    zip_ok, zip_evidence, entries = zip_audit(output)
    results: list[dict[str, object]] = []
    if not zip_ok:
        results.append(result("R001", 0.0, zip_evidence))
        for cid in list(WEIGHTS)[1:]:
            results.append(result(cid, 0.0, "核心 DOCX 无法安全解析，未继续计分"))
        return results, output

    try:
        document = Document(output)
        rows, selected, queue = source_truth(source)
        all_text = clean(" ".join(paragraph.text for paragraph in document.paragraphs) + " " + " ".join(cell.text for table in document.tables for row in table.rows for cell in row.cells))
        render_ok, render_evidence = render_audit(output)
        editable_core = len(document.tables) >= 13 and len(all_text) > 600
        r1 = input_digest == INPUT_SHA256 and only_expected and editable_core and render_ok
        results.append(result("R001", 1.0 if r1 else 0.0, f"{zip_evidence}；input_sha256={input_digest}；仅指定DOCX={only_expected}；原生结构={editable_core}；{render_evidence}"))

        r2_score, r2_evidence, actual_queue = criterion_queue_filter(document, all_text, selected, queue, rows)
        results.append(result("R002", r2_score, r2_evidence))
        r3_score, r3_evidence = criterion_queue_logic(actual_queue, queue)
        results.append(result("R003", r3_score, r3_evidence))
        r4_score, r4_evidence, r5_score, r5_evidence = criterion_profiles(document, queue)
        results.append(result("R004", r4_score, r4_evidence))
        results.append(result("R005", r5_score, r5_evidence))
        r6_score, r6_evidence = criterion_appendix(document, rows, selected)
        results.append(result("R006", r6_score, r6_evidence))
        r7_score, r7_evidence = criterion_native_objects(entries, queue)
        results.append(result("R007", r7_score, r7_evidence))
        r8_score, r8_evidence = criterion_structure(document, entries, render_ok, render_evidence, rows, selected)
        results.append(result("R008", r8_score, r8_evidence))
    except Exception as exc:
        results = [result("R001", 0.0, f"DOCX 或输入解析失败: {type(exc).__name__}: {exc}")]
        for cid in list(WEIGHTS)[1:]:
            results.append(result(cid, 0.0, "解析异常，未继续计分"))
    return results, output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--result", required=True, type=Path)
    parser.add_argument("--reward-file", required=True, type=Path)
    args = parser.parse_args()

    results, output = evaluate(args.artifact_dir)
    values = {item["criterion_id"]: float(item["score"]) for item in results}
    reward = sum(WEIGHTS[cid] * values.get(cid, 0.0) for cid in WEIGHTS)
    if values.get("R001", 0.0) < 1.0:
        reward = 0.0
    if values.get("R002", 0.0) < 1.0:
        reward = min(reward, 0.35)
    if values.get("R003", 0.0) < 1.0:
        reward = min(reward, 0.62)
    if values.get("R004", 0.0) < 1.0:
        reward = min(reward, 0.68)
    if values.get("R005", 0.0) < 1.0:
        reward = min(reward, 0.78)
    if values.get("R006", 0.0) < 1.0:
        reward = min(reward, 0.78)
    if values.get("R007", 0.0) < 1.0:
        reward = min(reward, 0.78)
    reward = round(reward, 6)
    passed = reward >= 0.82 and all(values.get(cid, 0.0) == 1.0 for cid in ("R001", "R002", "R003", "R004", "R005", "R006", "R007"))
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
