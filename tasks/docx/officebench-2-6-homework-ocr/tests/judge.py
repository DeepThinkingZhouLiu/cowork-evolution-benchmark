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


TASK_ID = "cowork-evo/officebench-2-6-homework-ocr"
INPUT_NAME = "homework.jpg"
OUTPUT_NAME = "homework.docx"
INPUT_SHA256 = "1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf"
RUBRIC_VERSION = "officebench-2-6-homework-ocr-v0.3"
JUDGE_VERSION = "officebench-2-6-homework-ocr-deterministic-v0.3"
WEIGHTS = {"R001": .10, "R002": .16, "R003": .18, "R004": .20, "R005": .12, "R006": .10, "R007": .08, "R008": .06}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "m": M_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def normalize_equation(value: object) -> str:
    """Treat Writer's OOXML multiplication glyph rewrite as semantically equivalent."""
    return re.sub(r"\s+", "", str(value if value is not None else "")).replace("∗", "*").replace("×", "*")


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def item(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def load_truth(root: Path) -> dict[str, object]:
    candidates = [root / "adaptation/transcription-truth.json", Path(__file__).with_name("transcription-truth.json")]
    for path in candidates:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    raise FileNotFoundError("transcription truth is unavailable")


def jpeg_size(path: Path) -> tuple[int, int]:
    data = path.read_bytes()
    if data[:2] != b"\xff\xd8":
        raise ValueError("input is not JPEG")
    index = 2
    while index + 9 < len(data):
        if data[index] != 0xFF:
            index += 1
            continue
        marker = data[index + 1]
        index += 2
        if marker in (0xD8, 0xD9):
            continue
        length = int.from_bytes(data[index:index + 2], "big")
        if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
            return int.from_bytes(data[index + 5:index + 7], "big"), int.from_bytes(data[index + 3:index + 5], "big")
        index += length
    raise ValueError("JPEG dimensions not found")


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, header: list[str]) -> list[list[list[str]]]:
    return [matrix for table in document.tables if (matrix := table_matrix(table)) and matrix[0] == header]


def exact_table(document: Document, header: list[str], expected: list[list[str]]) -> tuple[float, str]:
    matches = find_tables(document, header)
    if len(matches) != 1:
        return 0.0, f"table count={len(matches)} for {header}"
    actual = matches[0][1:]
    total = sum(len(row) for row in expected)
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = fraction(correct, total)
    if len(actual) != len(expected):
        score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"rows={len(actual)}/{len(expected)}; cells={correct}/{total}; exact={actual == expected}"


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
        with tempfile.TemporaryDirectory(prefix="judge-ocr-reconstruction-") as temp_dir:
            temp = Path(temp_dir)
            profile = (temp / "lo-profile").resolve()
            subprocess.run([
                "libreoffice", "--headless", f"-env:UserInstallation={profile.as_uri()}",
                "--convert-to", "pdf", "--outdir", str(temp), str(path),
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            pdf = temp / "homework.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            pages = int(re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE).group(1))
            width, height = map(float, re.search(r"^Page size:\s+([\d.]+) x ([\d.]+) pts", info, re.MULTILINE).groups())
            page_text = [clean(page) for page in subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30).split("\f") if clean(page)]
            contract = len(page_text) == 3 and "CSE 240A Homework 2" in page_text[0] and "So 7-stage processor is faster." in page_text[0] and "Transcription Evidence Ledger" in page_text[1] and "Pipeline Matrix Audit & Exception Closure" in page_text[2]
            return pages == 3 and width < height and contract, f"pages={pages}; portrait={width < height}; nonblank={len(page_text)}; ordered-contract={contract}"
    except Exception as exc:
        return False, f"LibreOffice render failed: {exc}"


def check_header_instructions(document: Document, truth: dict[str, object]):
    texts = [clean(paragraph.text) for paragraph in document.paragraphs]
    expected = [clean(truth["header"][key]) for key in ("title", "subtitle", "identity", "due")] + ["Instructions", *[clean(text) for text in truth["instructions"]], clean(truth["section_heading"])]
    counts = [texts.count(text) == 1 for text in expected]
    positions = [texts.index(text) if text in texts else -1 for text in expected]
    order_ok = all(a < b for a, b in zip(positions, positions[1:]))
    return (sum(counts) + int(order_ok)) / (len(expected) + 1), f"exact segments={sum(counts)}/{len(expected)}; source order={order_ok}; positions={positions}"


def check_questions(document: Document, entries: dict[str, bytes], truth: dict[str, object]):
    texts = [clean(paragraph.text) for paragraph in document.paragraphs]
    expected = []
    for question in truth["questions"]:
        expected.append(f"{question['number']}. {question['prompt']}")
        if "note" in question:
            expected.append(question["note"])
        if "conclusion" in question:
            expected.append(question["conclusion"])
    counts = []
    for text in expected:
        target = 2 if text == truth["questions"][0]["note"] else 1
        counts.append(texts.count(text) == target)
    root = etree.fromstring(entries["word/document.xml"])
    equations = [clean("".join(node.xpath(".//m:t/text()", namespaces=NS))) for node in root.xpath("//m:oMath", namespaces=NS)]
    equation_ok = [normalize_equation(value) for value in equations] == [normalize_equation(truth["questions"][2]["equation"])]
    return (sum(counts) + int(equation_ok)) / (len(counts) + 1), f"question/note/conclusion={sum(counts)}/{len(counts)}; native equation={equations}"


def expected_pipeline(truth: dict[str, object]) -> list[list[str]]:
    arrows = {
        ("lw R2, 0 (R5)", "5"): "W ↓",
        ("add R3, R2, R2", "8"): "W ↓",
        ("add R2, R2, R5", "9"): "M ↓",
    }
    rows = []
    for row in truth["pipeline"]["rows"]:
        instruction = row["instruction"]
        rows.append([instruction, *[arrows.get((instruction, cycle), row["stages"].get(cycle, "")) for cycle in truth["pipeline"]["cycles"]]])
    return rows


def segment_rows(truth: dict[str, object]) -> list[list[str]]:
    segments = [
        ("H1", "title", truth["header"]["title"], "top center"),
        ("H2", "subtitle", truth["header"]["subtitle"], "below title"),
        ("H3", "identity", truth["header"]["identity"], "header identity line"),
        ("H4", "due", truth["header"]["due"], "header due line"),
        ("S0", "label", "Instructions", "left heading"),
    ]
    segments.extend((f"I{index}", "instruction bullet", text, f"bullet {index}") for index, text in enumerate(truth["instructions"], 1))
    segments.append(("S1", "section heading", truth["section_heading"], "pipelining heading"))
    for question in truth["questions"]:
        segments.append((f"Q{question['number']}", "question prompt", f"{question['number']}. {question['prompt']}", f"visible prompt {question['number']}"))
        if "note" in question:
            segments.append((f"N{question['number']}", "question note", question["note"], f"note below prompt {question['number']}"))
        if "equation" in question:
            segments.append(("E3", "worked equation", question["equation"], "centered under Q3"))
            segments.append(("C3", "conclusion", question["conclusion"], "below equation"))
    return [[key, role, str(len(text)), hashlib.sha256(text.encode("utf-8")).hexdigest()[:12], clean(text[:42] + ("…" if len(text) > 42 else "")), locator] for key, role, text, locator in segments]


def check_evidence(document: Document, truth: dict[str, object]):
    source = truth["source"]
    anchor = [
        ["Source file", source["filename"], "Frozen runtime input"],
        ["SHA256", source["sha256"], "Exact package/environment bytes"],
        ["Raster geometry", f"{source['width_px']} × {source['height_px']} px", "JPEG; one visible page"],
        ["Reconstruction", "Page 1", "Editable text + native pipeline table"],
        ["Embedded source image", "0", "No screenshot or flattened-page fallback"],
    ]
    a_score, a_evidence = exact_table(document, ["Anchor", "Verified value", "Use"], anchor)
    s_score, s_evidence = exact_table(document, ["Key", "Role", "Chars", "SHA-12", "Text locator", "Scan location"], segment_rows(truth))
    return (a_score + s_score) / 2, f"anchor[{a_evidence}]; segments[{s_evidence}]"


def occupied_cycles(row: dict[str, object]) -> str:
    return " · ".join(f"{cycle}:{stage}" for cycle, stage in row["stages"].items())


def check_audit_tables(document: Document, truth: dict[str, object]):
    arrows = truth["pipeline"]["visible_down_arrows"]
    arrow_by_instruction = {row["instruction"]: "—" for row in truth["pipeline"]["rows"]}
    arrow_by_instruction["lw R2, 0 (R5)"] = arrows[0]
    arrow_by_instruction["add R3, R2, R2"] = arrows[1]
    arrow_by_instruction["add R2, R2, R5"] = arrows[2]
    pipeline = []
    for row in truth["pipeline"]["rows"]:
        stalls = "D@3–4" if row["instruction"] == "sub R10, R2, 10" else ("F@3–4" if row["instruction"] == "add R3, R2, R2" else "—")
        pipeline.append([row["instruction"], occupied_cycles(row), stalls, arrow_by_instruction[row["instruction"]]])
    conflict = [[entry["issue"], entry["source_evidence"], entry["treatment"]] for entry in truth["source_conflicts"]]
    closure = [
        ["Visible source pages", "1", "homework.jpg"], ["Editable reconstruction pages", "1", "Output page 1"],
        ["Question prompts", "3", "Preserved despite source count conflict"], ["Instruction bullets", "5", "Exact visible order"],
        ["Pipeline rows / cycles", "6 / 15", "Native 7 × 16 Word table"], ["Occupied stage cells", "34", "Includes repeated F/D stalls"],
        ["Visible dependency arrows", "3", "Explicit relations retained"], ["Native equation objects", "1", "Question 3 worked calculation"],
        ["Bookmarks / internal links", "4 / 3", "Section + Q1/Q2/Q3; links to questions"],
        ["Custom properties", "6", "SourceSHA256, Course, Homework, StudentName, StudentID, DueDate"],
    ]
    checks = [
        exact_table(document, ["Instruction", "Occupied cycle:stage cells", "Visible stall", "Down-arrow relation"], pipeline),
        exact_table(document, ["Issue", "Source evidence", "Required treatment"], conflict),
        exact_table(document, ["Token family", "Verified rendering", "Disambiguation"], truth["transcription_conventions"]),
        exact_table(document, ["Control", "Result", "Evidence / rule"], closure),
    ]
    return sum(score for score, _ in checks) / len(checks), "; ".join(evidence for _, evidence in checks)


def custom_properties(entries: dict[str, bytes]) -> dict[str, str]:
    if "docProps/custom.xml" not in entries:
        return {}
    root = etree.fromstring(entries["docProps/custom.xml"])
    result = {}
    for prop in root.xpath("//cp:property", namespaces=NS):
        children = list(prop)
        result[prop.get("name")] = clean(children[0].text if children else "")
    return result


def check_native(document: Document, entries: dict[str, bytes]):
    root = etree.fromstring(entries["word/document.xml"])
    expected_properties = {
        "SourceSHA256": INPUT_SHA256, "Course": "CSE 240A", "Homework": "Homework 2", "StudentName": "Ryan Wang", "StudentID": "A012345", "DueDate": "June 1, 2021",
    }
    actual_properties = custom_properties(entries)
    property_score = fraction(sum(actual_properties.get(key) == value for key, value in expected_properties.items()), len(expected_properties))
    bookmarks = root.xpath("//w:bookmarkStart/@w:name", namespaces=NS)
    links = root.xpath("//w:hyperlink/@w:anchor", namespaces=NS)
    nav_ok = bookmarks == ["question_1", "question_2", "question_3", "pipelining_section"] or bookmarks == ["pipelining_section", "question_1", "question_2", "question_3"]
    nav_ok = nav_ok and links == ["question_1", "question_2", "question_3"]
    core = document.core_properties
    core_ok = core.title == "CSE 240A Homework 2 Editable OCR Reconstruction" and core.category == "OCR reconstruction control"
    return (property_score * 6 + int(nav_ok) * 2 + int(core_ok) * 2) / 10, f"custom={actual_properties}; bookmarks={bookmarks}; links={links}; core={core_ok}"


def check_structure(document: Document, entries: dict[str, bytes], render_ok: bool, render_evidence: str):
    root = etree.fromstring(entries["word/document.xml"])
    headers = any(b"EDITABLE OCR RECONSTRUCTION" in payload and b"A012345" in payload for name, payload in entries.items() if name.startswith("word/header"))
    footer = any(b"PAGE" in payload and b"homework.jpg" in payload for name, payload in entries.items() if name.startswith("word/footer"))
    tables_ok = len(document.tables) == 7 and len(root.xpath("//w:tblHeader", namespaces=NS)) >= 7 and len(root.xpath("//w:cantSplit", namespaces=NS)) >= 58
    titles = [clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style and paragraph.style.name == "Title"]
    headings = [clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style and paragraph.style.name.startswith("Heading")]
    title_ok = titles == ["CSE 240A Homework 2"] and headings == ["Transcription Evidence Ledger", "Segment Fingerprints", "Pipeline Matrix Audit & Exception Closure", "Six-Instruction Pipeline Closure", "Source Conflicts — Preserve, Do Not Repair", "Verified Transcription Conventions", "Package Closure"]
    no_suspicious = not any(term in clean("\n".join(paragraph.text for paragraph in document.paragraphs)).casefold() for term in ("approved for submission", "grade awarded", "missing diagram reconstructed", "placeholder", "todo"))
    checks = [headers and footer, tables_ok, title_ok, no_suspicious, render_ok]
    return fraction(sum(checks), len(checks)), f"header/footer={headers and footer}; tables={tables_ok}; titles/headings={title_ok}; unsupported={not no_suspicious}; {render_evidence}"


def evaluate(artifact_dir: Path):
    task_root = Path(__file__).parents[1]
    input_path = task_root / "data/input_files" / INPUT_NAME
    output = artifact_dir / OUTPUT_NAME
    safe, safe_evidence, entries = ooxml_audit(output)
    if not safe:
        return [item(cid, 0, safe_evidence if cid == "R001" else "core DOCX unavailable") for cid in WEIGHTS], output
    try:
        truth = load_truth(task_root)
        digest = hashlib.sha256(input_path.read_bytes()).hexdigest()
        dimensions = jpeg_size(input_path)
        input_ok = digest == INPUT_SHA256 == truth["source"]["sha256"] and dimensions == (truth["source"]["width_px"], truth["source"]["height_px"])
        document = Document(output)
        render_ok, render_evidence = render(output)
        only_expected = sorted(path.name for path in artifact_dir.glob("*.docx")) == [OUTPUT_NAME]
        native = len(document.tables) == 7 and len(document.paragraphs) >= 27
        results = [item("R001", 1 if input_ok and only_expected and native and render_ok else 0, f"{safe_evidence}; input={digest}/{dimensions}; only_expected={only_expected}; native={native}; {render_evidence}")]
        r2, e2 = check_header_instructions(document, truth)
        results.append(item("R002", r2, e2))
        r3, e3 = check_questions(document, entries, truth)
        results.append(item("R003", r3, e3))
        r4, e4 = exact_table(document, ["Instruction", *truth["pipeline"]["cycles"]], expected_pipeline(truth))
        results.append(item("R004", r4, e4))
        r5, e5 = check_evidence(document, truth)
        results.append(item("R005", r5, e5))
        r6, e6 = check_audit_tables(document, truth)
        results.append(item("R006", r6, e6))
        r7, e7 = check_native(document, entries)
        results.append(item("R007", r7, e7))
        r8, e8 = check_structure(document, entries, render_ok, render_evidence)
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
    for cid, cap in (("R001", 0.0), ("R002", .30), ("R003", .42), ("R004", .55), ("R005", .68), ("R006", .76), ("R007", .80), ("R008", .82)):
        if values.get(cid, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    passed = reward >= .86 and all(values.get(cid, 0) == 1 for cid in WEIGHTS)
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
