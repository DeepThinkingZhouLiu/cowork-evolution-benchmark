from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree


TASK_ID = "cowork-evo/officebench-2-30-rename-student-id"
INPUT_NAME = "homework_withname.docx"
OUTPUT_NAME = "A12345678.docx"
INPUT_SHA256 = "7df2fd4bf45616e4e78f6adf4dfb5c10ac18886bc777bd43064740be12b8b03b"
BODY_SHA256 = "9215262c621e091ae431a8a2b35e56b5d75a92f8183693ce5f5c0a13ccaabf05"
RUBRIC_VERSION = "officebench-2-30-rename-student-id-v0.3"
JUDGE_VERSION = "officebench-2-30-rename-student-id-deterministic-v0.3"
WEIGHTS = {"R001": .10, "R002": .15, "R003": .22, "R004": .14, "R005": .13, "R006": .10, "R007": .09, "R008": .07}
ROLES = [
    "assignment title", "academic term", "student identity", "answering instruction", "grading policy", "submission instruction",
    "question 1", "question 1 option", "question 1 option", "question 1 option", "support instruction", "example answer", "quotation example",
]
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS = {"w": W_NS, "cp": CP_NS, "vt": VT_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def item(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, header: list[str]) -> list[list[list[str]]]:
    return [matrix for table in document.tables if (matrix := table_matrix(table)) and matrix[0] == header]


def exact_table(document: Document, header: list[str], expected: list[list[str]]) -> tuple[float, str]:
    tables = find_tables(document, header)
    if len(tables) != 1:
        return 0.0, f"table count={len(tables)} for {header}"
    actual = tables[0][1:]
    total = sum(len(row) for row in expected)
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = fraction(correct, total)
    if len(actual) != len(expected):
        score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"rows={len(actual)}/{len(expected)}; cells={correct}/{total}; exact={actual == expected}"


def paragraph_numbering(paragraph) -> tuple[int | None, int | None]:
    properties = paragraph._p.pPr
    numbering = properties.numPr if properties is not None else None
    if numbering is None or numbering.numId is None:
        return None, None
    return int(numbering.numId.val), int(numbering.ilvl.val if numbering.ilvl is not None else 0)


def numbering_label(paragraph) -> str:
    num_id, level = paragraph_numbering(paragraph)
    return "none" if num_id is None else f"numId {num_id} • level {level}"


def run_signature(run) -> tuple[object, ...]:
    size = run.font.size.pt if run.font.size is not None else None
    color = str(run.font.color.rgb) if run.font.color is not None and run.font.color.rgb is not None else None
    return (run.text, run.bold, run.italic, run.font.name, size, color)


def properties_signature(paragraph) -> bytes:
    properties = paragraph._p.pPr
    return etree.tostring(properties, method="c14n") if properties is not None else b""


def source_truth(source: Path):
    document = Document(source)
    paragraphs = list(document.paragraphs)
    texts = [clean(paragraph.text) for paragraph in paragraphs]
    identifiers = re.findall(r"\b[A-Z]\d{8}\b", "\n".join(texts))
    title = re.fullmatch(r"(CSE \d+): (Homework \d+)", texts[0])
    if len(paragraphs) != 13 or len(identifiers) != 1 or title is None:
        raise ValueError("frozen source shape or identity mismatch")
    return document, paragraphs, texts, {
        "student_id": identifiers[0],
        "course": title.group(1),
        "assignment": title.group(2),
        "term": texts[1],
        "body_sha256": hashlib.sha256("\n".join(texts).encode("utf-8")).hexdigest(),
    }


def ooxml_audit(path: Path):
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "DOCX missing or invalid OOXML", {}
    try:
        with zipfile.ZipFile(path) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        prohibited = [name for name in entries if "embeddings/" in name.lower() or name.lower().endswith((".bin", ".exe", ".docm", ".xlsm"))]
        external = []
        for name, payload in entries.items():
            if name.endswith(".rels"):
                root = etree.fromstring(payload)
                external.extend(root.xpath("//*[local-name()='Relationship'][@TargetMode='External']/@Target"))
        media = [name for name in entries if name.startswith("word/media/")]
        ok = not prohibited and not external and not media and "word/document.xml" in entries
        return ok, f"parts={len(entries)}; prohibited={prohibited}; external={external}; media={media}", entries
    except Exception as exc:
        return False, f"OOXML parse failed: {exc}", {}


def render(path: Path):
    try:
        with tempfile.TemporaryDirectory(prefix="judge-identity-control-") as temp_dir:
            temp = Path(temp_dir)
            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(temp), str(path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            pdf = temp / f"{path.stem}.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            pages = int(re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE).group(1))
            width, height = map(float, re.search(r"^Page size:\s+([\d.]+) x ([\d.]+) pts", info, re.MULTILINE).groups())
            page_text = subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30).split("\f")
            page_text = [clean(page) for page in page_text if clean(page)]
            page_contract = len(page_text) == 3 and "Submission Identity Control Sheet" in page_text[0] and "CSE 221: Homework 1" in page_text[1] and "Preservation Ledger & Closure" in page_text[2]
            return pages == 3 and width < height and page_contract, f"pages={pages}; portrait={width < height}; nonblank={len(page_text)}; ordered-page-contract={page_contract}"
    except Exception as exc:
        return False, f"LibreOffice render failed: {exc}"


def locate_original(document: Document, source_texts: list[str]):
    candidates = [index for index, paragraph in enumerate(document.paragraphs) if clean(paragraph.text) == source_texts[0]]
    if len(candidates) != 1:
        return None, [], f"assignment-title occurrences={candidates}"
    start = candidates[0]
    block = document.paragraphs[start:start + len(source_texts)]
    return start, block, f"original block start={start + 1}; paragraphs={len(block)}"


def check_identity(document: Document, output: Path, truth: dict[str, str]):
    identity = [
        ["Student ID", truth["student_id"], "Body paragraph 3 • run 2"],
        ["Course", truth["course"], "Body paragraph 1"],
        ["Assignment", truth["assignment"], "Body paragraph 1"],
        ["Academic term", truth["term"], "Body paragraph 2"],
        ["ID occurrence count", "1", "Whole original body"],
        ["ID pattern", "PASS", r"^[A-Z]\d{8}$"],
    ]
    score, evidence = exact_table(document, ["Field", "Extracted value", "Source / rule"], identity)
    filename_ok = output.name == f"{truth['student_id']}.docx"
    return (score * 18 + int(filename_ok) * 2) / 20, f"{evidence}; filename-derived={filename_ok}"


def check_content(block, source_paragraphs, source_texts):
    if len(block) != len(source_texts):
        return 0.0, f"block paragraphs={len(block)}/{len(source_texts)}"
    text_correct = sum(clean(actual.text) == expected for actual, expected in zip(block, source_texts))
    run_total = sum(len(paragraph.runs) for paragraph in source_paragraphs)
    run_correct = 0
    for actual, source in zip(block, source_paragraphs):
        run_correct += sum(a.text == b.text for a, b in zip(actual.runs, source.runs))
        if len(actual.runs) != len(source.runs):
            run_correct -= abs(len(actual.runs) - len(source.runs))
    run_correct = max(0, run_correct)
    score = (text_correct * 3 + min(run_correct, run_total)) / (len(source_texts) * 3 + run_total)
    return score, f"paragraph text={text_correct}/{len(source_texts)}; run text/boundaries={run_correct}/{run_total}"


def check_formatting(block, source_paragraphs, entries: dict[str, bytes], source_entries: dict[str, bytes]):
    if len(block) != len(source_paragraphs):
        return 0.0, "source block incomplete"
    run_total = sum(len(paragraph.runs) for paragraph in source_paragraphs)
    run_correct = sum(run_signature(a) == run_signature(b) for actual, source in zip(block, source_paragraphs) for a, b in zip(actual.runs, source.runs))
    paragraph_correct = sum(
        actual.style.name == source.style.name
        and paragraph_numbering(actual) == paragraph_numbering(source)
        and properties_signature(actual) == properties_signature(source)
        for actual, source in zip(block, source_paragraphs)
    )
    numbering_same = False
    if entries.get("word/numbering.xml") and source_entries.get("word/numbering.xml"):
        candidate_numbering = etree.fromstring(entries["word/numbering.xml"])
        source_numbering = etree.fromstring(source_entries["word/numbering.xml"])
        numbering_same = etree.tostring(candidate_numbering, method="c14n") == etree.tostring(source_numbering, method="c14n")
    score = (run_correct + paragraph_correct + int(numbering_same) * 4) / (run_total + len(source_paragraphs) + 4)
    return score, f"run formatting={run_correct}/{run_total}; paragraph style/pPr/numbering={paragraph_correct}/{len(source_paragraphs)}; numbering-part-preserved={numbering_same}"


def check_ledger(document: Document, source_paragraphs):
    expected = []
    for index, (paragraph, role) in enumerate(zip(source_paragraphs, ROLES), start=1):
        text = clean(paragraph.text)
        expected.append([str(index), role, numbering_label(paragraph), str(len(paragraph.runs)), str(len(text)), hashlib.sha256(text.encode("utf-8")).hexdigest()[:12], text[:34] + ("…" if len(text) > 34 else "")])
    return exact_table(document, ["#", "Role", "List state", "Runs", "Chars", "SHA-12", "Text locator"], expected)


def check_controls(document: Document, truth: dict[str, str]):
    filename = [
        ["Derived filename", f"{truth['student_id']}.docx", "PASS: stem equals extracted ID"],
        ["Extension", ".docx", "PASS: editable OOXML"],
        ["Expected deliverables", "1", "No sidecar or alternate-ID DOCX"],
        ["Source SHA256", INPUT_SHA256, "Frozen OfficeBench input"],
    ]
    filename_score, filename_evidence = exact_table(document, ["Control", "Value", "Decision / evidence"], filename)
    closure = [
        ["Original paragraphs", "13", "Input body"],
        ["Preserved assignment paragraphs", "13", "Contiguous page-2 block"],
        ["Paragraph text matches", "13 / 13", "Exact normalized text"],
        ["Run text sequence", "14 / 14", "Exact source run boundaries"],
        ["Native numbering", "4", "1 level 0 + 3 level 1"],
        ["Student ID occurrences", "1", "Original assignment block"],
        ["Source native objects", "0 / 0 / 0", "Tables / images / hyperlinks"],
        ["Identity bookmark / link", "1 / 1", "source_student_id"],
        ["Custom properties", "6", "StudentID, Course, Assignment, Term, SourceSHA256, PreservedParagraphs"],
        ["Body fingerprint", BODY_SHA256, "13 normalized paragraphs joined by LF"],
    ]
    closure_score, closure_evidence = exact_table(document, ["Control", "Result", "Evidence / rule"], closure)
    return (filename_score + closure_score) / 2, f"filename[{filename_evidence}]; closure[{closure_evidence}]"


def read_custom_properties(entries: dict[str, bytes]) -> dict[str, str]:
    if "docProps/custom.xml" not in entries:
        return {}
    root = etree.fromstring(entries["docProps/custom.xml"])
    output = {}
    for prop in root.xpath("//cp:property", namespaces=NS):
        children = list(prop)
        output[prop.get("name")] = clean(children[0].text if children else "")
    return output


def check_metadata_navigation(document: Document, entries: dict[str, bytes], truth: dict[str, str]):
    expected_properties = {
        "StudentID": truth["student_id"], "Course": truth["course"], "Assignment": truth["assignment"], "Term": truth["term"],
        "SourceSHA256": INPUT_SHA256, "PreservedParagraphs": "13",
    }
    properties = read_custom_properties(entries)
    custom_score = fraction(sum(properties.get(key) == value for key, value in expected_properties.items()), len(expected_properties))
    root = etree.fromstring(entries["word/document.xml"])
    bookmarks = root.xpath("//w:bookmarkStart/@w:name", namespaces=NS)
    links = [(node.get(f"{{{W_NS}}}anchor"), clean("".join(node.xpath(".//w:t/text()", namespaces=NS)))) for node in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS)]
    nav_ok = bookmarks == ["source_student_id"] and links == [("source_student_id", "Jump to the original Student ID")]
    core = document.core_properties
    core_ok = core.title == "CSE 221 Homework 1 Submission Control" and core.subject == "Identity-preserved submission package for A12345678" and core.category == "Submission identity control"
    return (custom_score * 6 + int(nav_ok) * 2 + int(core_ok) * 2) / 10, f"custom={properties}; navigation={links}/{bookmarks}; core={core_ok}"


def check_structure(document: Document, entries: dict[str, bytes], source_texts: list[str], start: int | None, render_ok: bool, render_evidence: str):
    root = etree.fromstring(entries["word/document.xml"])
    headers = any(b"SUBMISSION CONTROL" in payload and b"A12345678" in payload for name, payload in entries.items() if name.startswith("word/header"))
    page_field = any(b"PAGE" in payload for name, payload in entries.items() if name.startswith("word/footer"))
    tables_ok = len(document.tables) == 6 and len(root.xpath("//w:tblHeader", namespaces=NS)) >= 6 and len(root.xpath("//w:cantSplit", namespaces=NS)) >= 44
    titles = [clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style and paragraph.style.name == "Title"]
    title_ok = titles == ["Submission Identity Control Sheet", "Preservation Ledger & Closure"]
    nonblank = [clean(paragraph.text) for paragraph in document.paragraphs if clean(paragraph.text)]
    count_ok = len(nonblank) == 25
    between_ok = False
    if start is not None:
        appendix = [index for index, paragraph in enumerate(document.paragraphs) if clean(paragraph.text) == "Preservation Ledger & Closure"]
        if len(appendix) == 1:
            between = [clean(paragraph.text) for paragraph in document.paragraphs[start:appendix[0]] if clean(paragraph.text)]
            between_ok = between == source_texts
    checks = [headers and page_field, tables_ok, title_ok, count_ok and between_ok, render_ok]
    return fraction(sum(checks), len(checks)), f"header/page={headers and page_field}; native tables={tables_ok}; titles={titles}; nonblank={len(nonblank)}/25; untouched-block-boundary={between_ok}; {render_evidence}"


def evaluate(artifact_dir: Path):
    source = Path(__file__).parents[1] / "data/input_files" / INPUT_NAME
    output = artifact_dir / OUTPUT_NAME
    safe, safe_evidence, entries = ooxml_audit(output)
    if not safe:
        return [item(cid, 0, safe_evidence if cid == "R001" else "core DOCX unavailable") for cid in WEIGHTS], output
    try:
        source_document, source_paragraphs, source_texts, truth = source_truth(source)
        with zipfile.ZipFile(source) as archive:
            source_entries = {name: archive.read(name) for name in archive.namelist()}
        document = Document(output)
        render_ok, render_evidence = render(output)
        only_expected = sorted(path.name for path in artifact_dir.glob("*.docx")) == [OUTPUT_NAME]
        input_ok = hashlib.sha256(source.read_bytes()).hexdigest() == INPUT_SHA256 and truth["body_sha256"] == BODY_SHA256
        native = len(document.tables) == 6 and len(document.paragraphs) >= 24
        results = [item("R001", 1 if input_ok and only_expected and native and render_ok else 0, f"{safe_evidence}; input={input_ok}; only_expected={only_expected}; native={native}; {render_evidence}")]
        r2, e2 = check_identity(document, output, truth)
        results.append(item("R002", r2, e2))
        start, block, locate_evidence = locate_original(document, source_texts)
        r3, e3 = check_content(block, source_paragraphs, source_texts)
        results.append(item("R003", r3, f"{locate_evidence}; {e3}"))
        r4, e4 = check_formatting(block, source_paragraphs, entries, source_entries)
        results.append(item("R004", r4, e4))
        r5, e5 = check_ledger(document, source_paragraphs)
        results.append(item("R005", r5, e5))
        r6, e6 = check_controls(document, truth)
        results.append(item("R006", r6, e6))
        r7, e7 = check_metadata_navigation(document, entries, truth)
        results.append(item("R007", r7, e7))
        r8, e8 = check_structure(document, entries, source_texts, start, render_ok, render_evidence)
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
    for cid, cap in (("R001", 0.0), ("R002", .30), ("R003", .40), ("R004", .58), ("R005", .70), ("R006", .76), ("R007", .78), ("R008", .80)):
        if values.get(cid, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    passed = reward >= .84 and all(values.get(cid, 0) == 1 for cid in WEIGHTS)
    payload = {
        "task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results, "reward": reward, "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
