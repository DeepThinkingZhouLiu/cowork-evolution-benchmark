from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-2-18-highest-score-statement"
RUBRIC_VERSION = "officebench-2-18-highest-score-statement-v0.3"
JUDGE_VERSION = "officebench-2-18-highest-score-statement-deterministic-v0.3"
OUTPUT_NAME = "highest_score_recognition_review.docx"
WEIGHTS = {"R001": .08, "R002": .17, "R003": .12, "R004": .08, "R005": .11, "R006": .18, "R007": .08, "R008": .06, "R009": .07, "R010": .05}
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": WNS}


@dataclass(frozen=True)
class Record:
    source_row: int
    name: str
    midterm1: int
    midterm2: int

    @property
    def delta(self) -> int:
        return self.midterm2 - self.midterm1

    @property
    def direction(self) -> str:
        return "Down" if self.delta < 0 else ("Up" if self.delta > 0 else "No change")


def criterion(identifier: str, score: float, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def derive_truth(source_dir: Path) -> tuple[list[tuple[int, str, int]], list[Record]]:
    shortlist_book = load_workbook(source_dir / "nomination_shortlist.xlsx", data_only=True, read_only=True)
    shortlist_sheet = shortlist_book.active
    if [cell.value for cell in shortlist_sheet[1]] != ["Name", "midterm1"]:
        raise ValueError("nomination_shortlist.xlsx 表头不符合冻结证据")
    shortlist = [(number, str(name), int(score)) for number, (name, score) in enumerate(shortlist_sheet.iter_rows(min_row=2, values_only=True), start=2)]
    shortlist_book.close()
    register_book = load_workbook(source_dir / "detailed_midterm_register.xlsx", data_only=True, read_only=True)
    register_sheet = register_book.active
    if [cell.value for cell in register_sheet[1]] != ["Name", "midterm1", "midterm2"]:
        raise ValueError("detailed_midterm_register.xlsx 表头不符合冻结证据")
    records = [Record(number, str(name), int(first), int(second)) for number, (name, first, second) in enumerate(register_sheet.iter_rows(min_row=2, values_only=True), start=2)]
    register_book.close()
    if len(shortlist) != 4 or len(records) != 20:
        raise ValueError(f"冻结输入行数异常: shortlist={len(shortlist)}, register={len(records)}")
    return shortlist, records


def package_parts(path: Path) -> tuple[dict[str, bytes], etree._Element]:
    with zipfile.ZipFile(path) as package:
        damaged = package.testzip()
        if damaged:
            raise ValueError(f"CRC 损坏: {damaged}")
        parts = {name: package.read(name) for name in package.namelist()}
    return parts, etree.fromstring(parts["word/document.xml"])


def body_text(root: etree._Element) -> str:
    return "\n".join("".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS)) for paragraph in root.xpath("//w:body//w:p", namespaces=NS))


def table_by_header(document: Document, labels: list[str]):
    for table in document.tables:
        if not table.rows:
            continue
        values = [cell.text.strip() for cell in table.rows[0].cells]
        if values[:len(labels)] == labels:
            return table
    return None


def score_decision(text: str, shortlist: list[tuple[int, str, int]], records: list[Record]) -> tuple[float, str]:
    maximum = max(record.midterm1 for record in records)
    winners = [record for record in records if record.midterm1 == maximum]
    shortlist_pairs = {(name, score) for _, name, score in shortlist}
    omissions = [record for record in winners if (record.name, record.midterm1) not in shortlist_pairs]
    lower = text.lower()
    checks = [
        "hold" in lower,
        all(record.name in text and str(record.midterm1) in text for record in winners),
        all(record.name in text and ("omitted" in lower or "absent" in lower) for record in omissions),
        "co-highest" in lower or "tied for the highest" in lower,
        "release" in lower and ("do not" in lower or "not for release" in lower),
        not any(phrase in lower for phrase in ["is the sole highest", "is the only highest", "is the sole winner", "is the only winner", "achieved the sole highest"]),
    ]
    return sum(checks) / len(checks), f"max={maximum}, winners={[(r.source_row, r.name) for r in winners]}, omissions={[(r.source_row, r.name) for r in omissions]}, checks={checks}"


def score_drafts(document: Document, records: list[Record]) -> tuple[float, str]:
    maximum = max(record.midterm1 for record in records)
    winners = [record for record in records if record.midterm1 == maximum]
    paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    hits = 0
    details = []
    for record in winners:
        label_ok = any(record.name in text and "DRAFT" in text and "NOT FOR RELEASE" in text for text in paragraphs)
        body_ok = any(f"Dear {record.name}" in text and str(maximum) in text and ("tied" in text.lower() or "co-highest" in text.lower()) and not any(claim in text.lower() for claim in ["is the sole", "is the only"]) for text in paragraphs)
        hits += int(label_ok) + int(body_ok)
        details.append(f"{record.name}:label={label_ok},body={body_ok}")
    distinct = sum(1 for record in winners if any(f"Dear {record.name}" in text for text in paragraphs)) == len(winners)
    hits += int(distinct)
    return hits / (len(winners) * 2 + 1), "; ".join(details) + f"; distinct={distinct}"


def score_snapshot(document: Document, records: list[Record]) -> tuple[float, str]:
    table = table_by_header(document, ["Control metric", "Result", "Control metric", "Result"])
    if table is None:
        return 0.0, "未找到审计快照表"
    values = " | ".join(cell.text.strip() for row in table.rows[1:] for cell in row.cells)
    expected = [
        str(len(records)),
        str(len({record.name for record in records})),
        f"{max(r.midterm1 for r in records)} / {min(r.midterm1 for r in records)} / {max(r.midterm1 for r in records) - min(r.midterm1 for r in records)}",
        f"{max(r.midterm2 for r in records)} / {min(r.midterm2 for r in records)}",
    ]
    checks = [value in values for value in expected]
    return sum(checks) / len(checks), f"expected={expected}, checks={checks}"


def score_shortlist(document: Document, shortlist: list[tuple[int, str, int]], records: list[Record]) -> tuple[float, str]:
    headers = ["Shortlist row", "Name", "Listed M1", "Detailed exact match", "Detailed source row", "Control result"]
    table = table_by_header(document, headers)
    if table is None or len(table.rows) != len(shortlist) + 1:
        return 0.0, f"提名对账表行数异常: {len(table.rows) if table else 0}"
    hits = 0
    total = len(shortlist) * 6
    details = []
    for table_row, (source_row, name, score) in zip(table.rows[1:], shortlist):
        cells = [cell.text.strip() for cell in table_row.cells]
        matches = [record.source_row for record in records if record.name == name and record.midterm1 == score]
        checks = [cells[0] == str(source_row), cells[1] == name, cells[2] == str(score), cells[3] == ("Yes" if matches else "No"), cells[4] == ", ".join(map(str, matches)), ("Verified" in cells[5]) if matches else ("Unresolved" in cells[5])]
        hits += sum(checks)
        details.append(f"row{source_row}={sum(checks)}/6")
    return hits / total, ", ".join(details)


def score_ledger(document: Document, shortlist: list[tuple[int, str, int]], records: list[Record]) -> tuple[float, str]:
    headers = ["Source row", "Displayed name", "M1", "M2", "Delta", "Direction", "Exact shortlist pair", "Evidence / control note"]
    table = table_by_header(document, headers)
    if table is None or len(table.rows) != len(records) + 1:
        return 0.0, f"详细台账行数异常: {len(table.rows) if table else 0}"
    shortlist_pairs = {(name, score) for _, name, score in shortlist}
    maximum = max(record.midterm1 for record in records)
    winners = [record for record in records if record.midterm1 == maximum]
    omissions = [record for record in winners if (record.name, record.midterm1) not in shortlist_pairs]
    hits = 0
    total = len(records) * 8
    weak_rows = []
    for table_row, record in zip(table.rows[1:], records):
        cells = [cell.text.strip() for cell in table_row.cells]
        exact = (record.name, record.midterm1) in shortlist_pairs
        note_ok = ("OMITTED" in cells[7] if record in omissions else ("Co-highest" in cells[7] if record in winners else ("listed" in cells[7].lower() if exact else "Not listed" in cells[7])))
        checks = [cells[0] == str(record.source_row), cells[1] == record.name, cells[2] == str(record.midterm1), cells[3] == str(record.midterm2), cells[4] == f"{record.delta:+d}", cells[5] == record.direction, cells[6] == ("Yes" if exact else "No"), note_ok]
        hits += sum(checks)
        if sum(checks) < len(checks):
            weak_rows.append(f"{record.source_row}:{sum(checks)}/8")
    return hits / total, f"rows={len(table.rows) - 1}/20, checks={hits}/{total}, weak={weak_rows}"


def score_comments(parts: dict[str, bytes], root: etree._Element, records: list[Record], shortlist: list[tuple[int, str, int]]) -> tuple[float, str]:
    if "word/comments.xml" not in parts:
        return 0.0, "无 Word 原生批注部件"
    comments = etree.fromstring(parts["word/comments.xml"])
    notes = ["".join(node.text or "" for node in comment.xpath(".//w:t", namespaces=NS)) for comment in comments.xpath("//w:comment", namespaces=NS)]
    maximum = max(record.midterm1 for record in records)
    winners = [record for record in records if record.midterm1 == maximum]
    shortlist_pairs = {(name, score): row for row, name, score in shortlist}
    starts = root.xpath("//w:commentRangeStart", namespaces=NS)
    hits = 0
    total = len(winners) * 2
    details = []
    for record in winners:
        anchor_ok = False
        for start in starts:
            paragraph = start.getparent()
            paragraph_text = "".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS))
            if record.name in paragraph_text and "DRAFT" in paragraph_text and "NOT FOR RELEASE" in paragraph_text:
                anchor_ok = True
        required = [record.name, str(record.source_row), str(record.midterm1), "detailed_midterm_register.xlsx"]
        shortlist_row = shortlist_pairs.get((record.name, record.midterm1))
        if shortlist_row:
            note_ok = any(all(value in note for value in required) and str(shortlist_row) in note and "nomination_shortlist.xlsx" in note for note in notes)
        else:
            note_ok = any(all(value in note for value in required) and ("No" in note or "omit" in note.lower()) and "nomination_shortlist.xlsx" in note for note in notes)
        hits += int(anchor_ok) + int(note_ok)
        details.append(f"{record.name}:anchor={anchor_ok},note={note_ok}")
    return hits / total, f"comments={len(notes)}, " + "; ".join(details)


def score_navigation(root: etree._Element, records: list[Record]) -> tuple[float, str]:
    bookmarks = {node.get(f"{{{WNS}}}name", "") for node in root.xpath("//w:bookmarkStart", namespaces=NS)}
    anchors = [node.get(f"{{{WNS}}}anchor", "") for node in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS)]
    maximum = max(record.midterm1 for record in records)
    winners = [record for record in records if record.midterm1 == maximum]
    required = {"executive_finding", "shortlist_reconciliation", "controlled_drafts", "record_ledger"} | {f"draft_{record.name.lower()}" for record in winners} | {f"record_{record.source_row}" for record in records}
    instructions = [node.text or "" for node in root.xpath("//w:instrText", namespaces=NS)] + list(root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
    toc = "TOC" in " ".join(map(str, instructions)).upper()
    checks = [required <= bookmarks, all(anchor in bookmarks for anchor in anchors), len(anchors) >= 6, toc]
    return sum(checks) / len(checks), f"bookmarks={len(bookmarks)}, required={len(required)}, links={len(anchors)}, toc={toc}, checks={checks}"


def score_structure(parts: dict[str, bytes], root: etree._Element, document: Document) -> tuple[float, str]:
    sections = root.xpath("//w:sectPr", namespaces=NS)
    landscape = any(section.xpath("./w:pgSz[@w:orient='landscape']", namespaces=NS) for section in sections)
    headers = " ".join("".join(node.text or "" for node in etree.fromstring(payload).xpath("//w:t", namespaces=NS)) for name, payload in parts.items() if name.startswith("word/header") and name.endswith(".xml"))
    footer_instructions = []
    for name, payload in parts.items():
        if name.startswith("word/footer") and name.endswith(".xml"):
            footer_root = etree.fromstring(payload)
            footer_instructions += [node.text or "" for node in footer_root.xpath("//w:instrText", namespaces=NS)]
            footer_instructions += list(footer_root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
    fields = " ".join(map(str, footer_instructions)).upper()
    ledger_tables = root.xpath("//w:tbl[w:tr[1]//w:t[contains(., 'Source row')]]", namespaces=NS)
    repeat_and_no_split = False
    if ledger_tables:
        rows = ledger_tables[0].xpath("./w:tr", namespaces=NS)
        repeat_and_no_split = bool(rows and rows[0].xpath("./w:trPr/w:tblHeader", namespaces=NS) and all(row.xpath("./w:trPr/w:cantSplit", namespaces=NS) for row in rows))
    properties = document.core_properties
    checks = [len(sections) >= 2, landscape, "HOLD" in headers and "MIDTERM" in headers, "PAGE" in fields and "NUMPAGES" in fields, repeat_and_no_split, "Release-Readiness" in (properties.title or "") and (properties.author or "") == "Academic Records Review"]
    return sum(checks) / len(checks), f"sections={len(sections)}, landscape={landscape}, header={headers!r}, footer_fields={fields!r}, repeat_no_split={repeat_and_no_split}, title={properties.title!r}, author={properties.author!r}"


def score_boundary(parts: dict[str, bytes], text: str, document: Document, records: list[Record]) -> tuple[float, str]:
    relationships = []
    for name, payload in parts.items():
        if name.endswith(".rels"):
            root = etree.fromstring(payload)
            relationships.extend(root.findall(f"{{{RELNS}}}Relationship"))
    external = any(relationship.get("TargetMode") == "External" for relationship in relationships)
    unsafe = any(name.endswith("vbaProject.bin") or "oleObject" in name or name.endswith("afchunk.mht") for name in parts)
    forbidden_patterns = [
        r"student id\s*[:#]?\s*\d+",
        r"course code\s*[:#]\s*[a-z]{2,}\d+",
        r"gpa\s*(?:is|=|:)?\s*[0-4](?:\.\d+)?",
        r"letter grade\s*(?:is|=|:)?\s*[a-f][+-]?\b",
        r"professor\s+[a-z]+",
        r"(?:is|are) eligible for (?:the )?award",
    ]
    forbidden = [pattern for pattern in forbidden_patterns if re.search(pattern, text.lower())]
    counts = Counter(record.name for record in records)
    duplicates = {name: count for name, count in counts.items() if count > 1}
    table = table_by_header(document, ["Displayed name", "Record count", "Required treatment"])
    duplicate_hits = 0
    if table is not None:
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows[1:]]
        for name, count in duplicates.items():
            duplicate_hits += int(any(cells[0] == name and cells[1] == str(count) and "separate" in cells[2].lower() and "no student id" in cells[2].lower() for cells in rows))
    checks = [duplicate_hits == len(duplicates), "cannot be merged" in text.lower() or "cannot be safely merged" in text.lower(), "obtain a stable student identifier" in text.lower(), not external, not unsafe, not forbidden, len(document.tables) >= 5 and len(text) >= 2000]
    return sum(checks) / len(checks), f"duplicate_checks={duplicate_hits}/{len(duplicates)}, external={external}, unsafe={unsafe}, forbidden={forbidden}, tables={len(document.tables)}, text_chars={len(text)}, checks={checks}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--result", required=True, type=Path)
    parser.add_argument("--reward-file", required=True, type=Path)
    arguments = parser.parse_args()
    source_dir = Path(__file__).resolve().parents[1] / "data/input_files"
    output = arguments.artifact_dir / OUTPUT_NAME
    artifacts = [path for path in arguments.artifact_dir.iterdir() if not path.name.startswith("~$")] if arguments.artifact_dir.is_dir() else []
    docx_files = [path for path in artifacts if path.suffix.lower() == ".docx"]
    results = []
    try:
        shortlist, records = derive_truth(source_dir)
        if len(artifacts) != 1 or len(docx_files) != 1 or docx_files[0].name != OUTPUT_NAME or not output.is_file() or not zipfile.is_zipfile(output):
            raise ValueError(f"交付文件={[path.name for path in artifacts]}")
        parts, root = package_parts(output)
        required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}
        if not required <= parts.keys():
            raise ValueError("缺少 DOCX 核心部件")
        document = Document(output)
        text = body_text(root)
        results.append(criterion("R001", 1.0, f"唯一交付={OUTPUT_NAME}，OOXML 可解析，members={len(parts)}"))
        score, evidence = score_decision(text, shortlist, records); results.append(criterion("R002", score, evidence))
        score, evidence = score_drafts(document, records); results.append(criterion("R003", score, evidence))
        score, evidence = score_snapshot(document, records); results.append(criterion("R004", score, evidence))
        score, evidence = score_shortlist(document, shortlist, records); results.append(criterion("R005", score, evidence))
        score, evidence = score_ledger(document, shortlist, records); results.append(criterion("R006", score, evidence))
        score, evidence = score_comments(parts, root, records, shortlist); results.append(criterion("R007", score, evidence))
        score, evidence = score_navigation(root, records); results.append(criterion("R008", score, evidence))
        score, evidence = score_structure(parts, root, document); results.append(criterion("R009", score, evidence))
        score, evidence = score_boundary(parts, text, document, records); results.append(criterion("R010", score, evidence))
    except Exception as exc:
        results = [criterion("R001", 0.0, f"读取或唯一文件门槛失败: {exc}")]
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[identifier] * values.get(identifier, 0.0) for identifier in WEIGHTS)
    hurdle_failures = []
    if values.get("R001", 0.0) < 1.0:
        hurdle_failures.append("R001")
        reward = 0.0
    caps = {"R002": .45, "R003": .55, "R004": .79, "R005": .72, "R006": .72, "R007": .78, "R009": .79, "R010": .75}
    for identifier, cap in caps.items():
        if values.get(identifier, 0.0) < 1.0:
            hurdle_failures.append(identifier)
            reward = min(reward, cap)
    reward = round(reward, 6)
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= .8 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
    }
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    arguments.reward_file.parent.mkdir(parents=True, exist_ok=True)
    arguments.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
