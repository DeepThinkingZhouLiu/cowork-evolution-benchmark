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


TASK_ID = "cowork-evo/officebench-3-22-homework-folder"
RUBRIC_VERSION = "officebench-3-22-homework-folder-v0.3"
JUDGE_VERSION = "officebench-3-22-homework-folder-deterministic-v0.3"
INPUT_SHA256 = "1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf"
MASTER_PATH = Path("homework.docx")
REVIEW_PATH = Path("homework/homework.docx")
WEIGHTS = {"R001": .12, "R002": .16, "R003": .17, "R004": .13, "R005": .12, "R006": .12, "R007": .10, "R008": .08}
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "m": M_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def normalize_equation(value: object) -> str:
    return re.sub(r"\s+", "", str(value if value is not None else "")).replace("∗", "*").replace("×", "*")


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def item(cid: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": cid, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def load_truth(root: Path) -> dict[str, object]:
    for path in (root / "adaptation/package-truth.json", Path(__file__).with_name("package-truth.json")):
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    raise FileNotFoundError("package truth is unavailable")


def canonical_payload(truth: dict[str, object]) -> dict[str, object]:
    return {key: truth[key] for key in ("header", "instructions", "section_heading", "questions", "pipeline")}


def canonical_digest(truth: dict[str, object]) -> str:
    payload = json.dumps(canonical_payload(truth), ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


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


def ooxml_audit(path: Path):
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "missing or invalid DOCX", {}
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


def render(path: Path, role: str):
    try:
        with tempfile.TemporaryDirectory(prefix=f"judge-controlled-{role}-") as temp_dir:
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
            texts = [clean(page) for page in subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, timeout=30).split("\f") if clean(page)]
            expected = [
                ["Controlled Digitization Master", "Canonical Source Payload", "Canonical Payload Register", "Master Handoff & Pair Integrity"],
                ["OCR Review Working Copy", "Canonical Source Payload", "Ambiguity Review Queue", "Reviewer Handoff & Synchronization"],
            ][role == "review"]
            ordered = len(texts) == 4 and all(expected[index] in texts[index] for index in range(4))
            ok = pages == 4 and width < height and ordered
            return ok, f"pages={pages}; portrait={width < height}; nonblank={len(texts)}; ordered={ordered}"
    except Exception as exc:
        return False, f"render failed: {type(exc).__name__}: {exc}"


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, header: list[str]) -> list[list[list[str]]]:
    return [matrix for table in document.tables if (matrix := table_matrix(table)) and matrix[0] == header]


def exact_table(document: Document, header: list[str], expected: list[list[str]]) -> tuple[float, str]:
    matches = find_tables(document, header)
    if len(matches) != 1:
        return 0.0, f"{header}: count={len(matches)}"
    actual = matches[0][1:]
    expected = [[clean(value) for value in row] for row in expected]
    total = sum(len(row) for row in expected)
    correct = sum(a == b for row, target in zip(actual, expected) for a, b in zip(row, target))
    score = fraction(correct, total)
    if len(actual) != len(expected):
        score *= fraction(min(len(actual), len(expected)), max(len(actual), len(expected)))
    return score, f"{header}: rows={len(actual)}/{len(expected)}; cells={correct}/{total}; exact={actual == expected}"


def expected_pipeline(truth: dict[str, object]) -> list[list[str]]:
    arrows = {
        ("lw R2, 0 (R5)", "5"): "W ↓",
        ("add R3, R2, R2", "8"): "W ↓",
        ("add R2, R2, R5", "9"): "M ↓",
    }
    output = []
    for row in truth["pipeline"]["rows"]:
        instruction = row["instruction"]
        output.append([instruction, *[arrows.get((instruction, cycle), row["stages"].get(cycle, "")) for cycle in truth["pipeline"]["cycles"]]])
    return output


def canonical_blocks(truth: dict[str, object]) -> list[list[str]]:
    blocks: list[tuple[str, str, str]] = [
        ("H1", "title", truth["header"]["title"]), ("H2", "subtitle", truth["header"]["subtitle"]),
        ("H3", "identity", truth["header"]["identity"]), ("H4", "due", truth["header"]["due"]),
        ("S0", "label", "Instructions"),
    ]
    blocks.extend((f"I{index}", "instruction", text) for index, text in enumerate(truth["instructions"], 1))
    blocks.append(("S1", "section", truth["section_heading"]))
    for entry in truth["questions"]:
        blocks.append((f"Q{entry['number']}", "prompt", f"{entry['number']}. {entry['prompt']}"))
        if "note" in entry:
            blocks.append((f"N{entry['number']}", "note", entry["note"]))
        if "equation" in entry:
            blocks.append(("E3", "equation", entry["equation"]))
            blocks.append(("C3", "conclusion", entry["conclusion"]))
    return [[key, role, str(len(text)), hashlib.sha256(text.encode("utf-8")).hexdigest()[:12], clean(text[:58] + ("…" if len(text) > 58 else ""))] for key, role, text in blocks]


def custom_properties(entries: dict[str, bytes]) -> dict[str, str]:
    payload = entries.get("docProps/custom.xml")
    if not payload:
        return {}
    root = etree.fromstring(payload)
    return {node.get("name"): clean("".join(node.xpath(".//text()"))) for node in root.xpath("//cp:property", namespaces=NS)}


def source_text_score(document: Document, truth: dict[str, object]) -> tuple[float, str]:
    texts = [clean(paragraph.text) for paragraph in document.paragraphs if clean(paragraph.text)]
    expected = [truth["header"][key] for key in ("title", "subtitle", "identity", "due")]
    expected.extend(["Instructions", *truth["instructions"], truth["section_heading"]])
    expected.extend([
        f"1. {truth['questions'][0]['prompt']}", truth["questions"][0]["note"],
        f"2. {truth['questions'][1]['prompt']}", truth["questions"][1]["note"],
        f"3. {truth['questions'][2]['prompt']}", truth["questions"][2]["conclusion"],
    ])
    expected = [clean(value) for value in expected]
    unique = list(dict.fromkeys(expected))
    exact = sum(texts.count(value) == expected.count(value) for value in unique)
    cursor = 0
    ordered = True
    for value in expected:
        try:
            cursor = texts.index(value, cursor) + 1
        except ValueError:
            ordered = False
            break
    unsupported = not any(term in clean(" ".join(texts)).casefold() for term in ("diagram reconstructed", "approved for release", "answer key", "grade awarded"))
    return (exact + int(ordered) + int(unsupported)) / (len(unique) + 2), f"exact={exact}/{len(unique)}; ordered={ordered}; unsupported_claim={not unsupported}"


def equation_score(entries: dict[str, bytes], truth: dict[str, object]) -> tuple[float, str]:
    root = etree.fromstring(entries["word/document.xml"])
    values = [clean("".join(node.xpath(".//m:t/text()", namespaces=NS))) for node in root.xpath("//m:oMath", namespaces=NS)]
    expected = truth["questions"][2]["equation"]
    return (1.0 if [normalize_equation(value) for value in values] == [normalize_equation(expected)] else 0.0), f"native equations={values}"


def check_pair(master: Document, review: Document, master_entries: dict[str, bytes], review_entries: dict[str, bytes], truth: dict[str, object], master_path: Path, review_path: Path):
    digest = canonical_digest(truth)
    expected_common = {
        "PackageID": truth["package"]["package_id"], "PackageVersion": truth["package"]["version"],
        "SourceSHA256": INPUT_SHA256, "CanonicalPayloadSHA256": digest, "PairCount": "2",
        "ReleaseState": truth["package"]["release_state"],
    }
    master_props = custom_properties(master_entries)
    review_props = custom_properties(review_entries)
    checks = [master_path.read_bytes() != review_path.read_bytes()]
    for key, value in expected_common.items():
        checks.extend((master_props.get(key) == value, review_props.get(key) == value))
    checks.extend((master_props.get("DocumentRole") == "MASTER", master_props.get("CopyPath") == str(MASTER_PATH), review_props.get("DocumentRole") == "REVIEW", review_props.get("CopyPath") == str(REVIEW_PATH)))
    return fraction(sum(checks), len(checks)), f"checks={sum(checks)}/{len(checks)}; master={master_props}; review={review_props}"


def check_master(document: Document, truth: dict[str, object]):
    digest = canonical_digest(truth)
    scores = []
    scores.append(exact_table(document, ["Block", "Role", "Chars", "SHA-12", "Verified locator"], canonical_blocks(truth)))
    scores.append(exact_table(document, ["Object", "Count", "Editability", "Scope"], [
        ["Pipeline matrix", "1", "Native 7 × 16 table", "Both copies"], ["Worked equation", "1", "Native OMML", "Both copies"],
        ["Source image", "0", "Not embedded or flattened", "Both copies"], ["Bookmarks / links", "3 / 3", "Internal navigation", "Each copy"],
        ["Review comments", "6", "Source-bound decisions", "Review copy only"], ["Review controls", "4", "Unlocked rich-text controls", "Review copy only"],
    ]))
    scores.append(exact_table(document, ["Checkpoint", "Master", "Review", "Required relation"], [
        ["Package ID", truth["package"]["package_id"], truth["package"]["package_id"], "Equal"],
        ["Source SHA256", INPUT_SHA256[:16], INPUT_SHA256[:16], "Equal; exact full property"],
        ["Canonical payload SHA256", digest[:16], digest[:16], "Equal; exact full property"],
        ["Canonical text blocks", "18", "18", "Exact and ordered"], ["Pipeline rows / cycles", "6 / 15", "6 / 15", "Exact cells and arrows"],
        ["Document role", "MASTER", "REVIEW", "Must differ"], ["Role-specific controls", "None", "4 controls + 6 comments", "Must differ"],
    ]))
    scores.append(exact_table(document, ["ID", "Evidence", "Master treatment"], [
        ["D01", truth["review_issues"][0][2], truth["review_issues"][0][3]], ["D02", truth["review_issues"][1][2], truth["review_issues"][1][3]],
    ]))
    scores.append(exact_table(document, ["Control", "Result", "Evidence / rule"], [
        ["Required DOCX paths", "2", "Root master + nested review copy"], ["Byte-identical copies", "No", "Role layers intentionally differ"],
        ["Canonical payload synchronized", "Yes", digest], ["Embedded media", "0", "Editable Word objects only"],
        ["Source conflicts disclosed", "2", "No repair or unsupported inference"], ["Review queue items", "6", "Resolved only in working copy workflow"],
        ["External relationships", "0", "Offline-safe package"], ["Release authority", "Master", "Review copy cannot self-approve"],
    ]))
    return sum(score for score, _ in scores) / len(scores), " | ".join(evidence for _, evidence in scores)


def check_review_issues(document: Document, entries: dict[str, bytes], truth: dict[str, object]):
    expected = [list(issue) for issue in truth["review_issues"]]
    table_score, table_evidence = exact_table(document, ["ID", "Issue", "Source evidence", "Required disposition"], expected)
    payload = entries.get("word/comments.xml", b"")
    comments = []
    if payload:
        root = etree.fromstring(payload)
        for node in root.xpath("//w:comment", namespaces=NS):
            comments.append((node.get(f"{{{W_NS}}}author"), node.get(f"{{{W_NS}}}initials"), clean("".join(node.xpath(".//w:t/text()", namespaces=NS)))))
    expected_comments = [("OCR QA Control", "OQC", f"{issue[0]} control: {issue[3]}") for issue in truth["review_issues"]]
    document_root = etree.fromstring(entries["word/document.xml"])
    ranges = len(document_root.xpath("//w:commentRangeStart", namespaces=NS))
    references = len(document_root.xpath("//w:commentReference", namespaces=NS))
    comment_score = fraction(sum((actual == target) for actual, target in zip(comments, expected_comments)) + int(ranges == 6) + int(references == 6), len(expected_comments) + 2)
    if len(comments) != len(expected_comments):
        comment_score *= fraction(min(len(comments), len(expected_comments)), max(len(comments), len(expected_comments)))
    return (table_score + comment_score) / 2, f"{table_evidence}; comments={len(comments)}/6; exact={comments == expected_comments}; ranges={ranges}; refs={references}"


def check_controls(master_entries: dict[str, bytes], review_entries: dict[str, bytes], truth: dict[str, object]):
    def controls(entries: dict[str, bytes]):
        root = etree.fromstring(entries["word/document.xml"])
        values = []
        for node in root.xpath("//w:sdt", namespaces=NS):
            tag = node.xpath("string(w:sdtPr/w:tag/@w:val)", namespaces=NS)
            alias = node.xpath("string(w:sdtPr/w:alias/@w:val)", namespaces=NS)
            value = clean("".join(node.xpath("w:sdtContent//w:t/text()", namespaces=NS)))
            locked = bool(node.xpath("w:sdtPr/w:lock", namespaces=NS))
            values.append((tag, alias, value, locked))
        return values
    master = controls(master_entries)
    review = controls(review_entries)
    expected = [(tag, alias, value, False) for tag, alias, value in truth["package"]["review_controls"]]
    correct = sum(actual == target for actual, target in zip(review, expected)) + int(master == [])
    total = len(expected) + 1
    if len(review) != len(expected):
        correct = max(0, correct - abs(len(review) - len(expected)))
    return fraction(correct, total), f"master={master}; review={review}; expected={expected}"


def check_structure(document: Document, entries: dict[str, bytes], role: str, render_ok: bool, render_evidence: str):
    root = etree.fromstring(entries["word/document.xml"])
    bookmarks = root.xpath("//w:bookmarkStart/@w:name", namespaces=NS)
    links = root.xpath("//w:hyperlink/@w:anchor", namespaces=NS)
    expected_bookmarks = ["canonical_payload", "payload_register", "master_handoff"] if role == "master" else ["canonical_payload", "review_queue", "review_handoff"]
    expected_headings = [
        "Package Control", "Copy Role Matrix", "Release Boundary", "Canonical Source Payload", "Canonical Payload Register", "Native Object Register", "Master Handoff & Pair Integrity", "Cross-Copy Contract", "Source Exceptions", "Master Closure",
    ] if role == "master" else [
        "Review Controls", "Copy Role Matrix", "QA Boundary", "Canonical Source Payload", "Ambiguity Review Queue", "Comment Contract", "Reviewer Handoff & Synchronization", "Control Inventory", "Cross-Copy Checklist", "Reviewer Boundary",
    ]
    headings = [clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style and paragraph.style.name.startswith("Heading")]
    header_term = b"CONTROLLED MASTER" if role == "master" else b"QA WORKING COPY"
    header_ok = any(header_term in payload and b"OCR-HW2-1185885B" in payload for name, payload in entries.items() if name.startswith("word/header"))
    footer_ok = any(b"OCR-HW2-1185885B" in payload and b"PAGE" in payload for name, payload in entries.items() if name.startswith("word/footer"))
    table_headers = len(root.xpath("//w:tblHeader", namespaces=NS))
    table_rows = len(root.xpath("//w:tr", namespaces=NS))
    no_splits = len(root.xpath("//w:cantSplit", namespaces=NS))
    no_unsupported = not any(term in clean(" ".join(paragraph.text for paragraph in document.paragraphs)).casefold() for term in ("approved for release", "diagram reconstructed", "placeholder", "todo"))
    body = [clean(paragraph.text) for paragraph in document.paragraphs]
    role_content = [
        "The master publishes only source-visible content plus audit metadata. The paired review copy carries unresolved decisions. Neither copy invents the absent Question 2 diagram, answers, grades, or institutional policy.",
    ] if role == "master" else [
        "Verify against the frozen source before changing a QA control.",
        "Do not edit the canonical transcript, pipeline cells, equation, or issue evidence.",
        "Do not mark the copy ready for release while D01 or D02 remains unresolved by source evidence.",
        "Return the reviewed working copy alongside the unchanged controlled master.",
    ]
    role_content_ok = all(value in body for value in role_content)
    checks = [bookmarks == expected_bookmarks, links == expected_bookmarks, headings == expected_headings, header_ok, footer_ok, table_headers == len(document.tables), no_splits == table_rows, no_unsupported, role_content_ok, render_ok]
    return fraction(sum(checks), len(checks)), f"checks={sum(checks)}/{len(checks)}; bookmarks={bookmarks}; links={links}; headings={headings}; headers={table_headers}/{len(document.tables)}; no-split={no_splits}/{table_rows}; unsupported={not no_unsupported}; {render_evidence}"


def evaluate(artifact_dir: Path):
    task_root = Path(__file__).parents[1]
    truth = load_truth(task_root)
    source = task_root / "data/input_files/homework.jpg"
    master_path = artifact_dir / MASTER_PATH
    review_path = artifact_dir / REVIEW_PATH
    master_safe, master_evidence, master_entries = ooxml_audit(master_path)
    review_safe, review_evidence, review_entries = ooxml_audit(review_path)
    expected_files = sorted(str(path.relative_to(artifact_dir)) for path in artifact_dir.rglob("*.docx")) if artifact_dir.is_dir() else []
    input_digest = hashlib.sha256(source.read_bytes()).hexdigest() if source.is_file() else "missing"
    try:
        input_dimensions = jpeg_size(source)
    except Exception:
        input_dimensions = (0, 0)
    if not master_safe or not review_safe:
        evidence = f"master[{master_evidence}]; review[{review_evidence}]"
        return [item(cid, 0, evidence if cid == "R001" else "complete two-copy package unavailable") for cid in WEIGHTS], [master_path, review_path]
    try:
        master = Document(master_path)
        review = Document(review_path)
        master_render, master_render_evidence = render(master_path, "master")
        review_render, review_render_evidence = render(review_path, "review")
        input_ok = input_digest == INPUT_SHA256 == truth["source"]["sha256"] and input_dimensions == (1700, 2200)
        files_ok = expected_files == [str(MASTER_PATH), str(REVIEW_PATH)]
        native = len(master.tables) == 8 and len(review.tables) == 7
        results = [item("R001", 1 if input_ok and files_ok and native and master_render and review_render else 0, f"master[{master_evidence}]; review[{review_evidence}]; input={input_digest}/{input_dimensions}; files={expected_files}; tables={len(master.tables)}/{len(review.tables)}; render={master_render_evidence} | {review_render_evidence}")]
        m_text, m_text_evidence = source_text_score(master, truth)
        r_text, r_text_evidence = source_text_score(review, truth)
        results.append(item("R002", (m_text + r_text) / 2, f"master[{m_text_evidence}]; review[{r_text_evidence}]"))
        m_pipeline, m_pipeline_evidence = exact_table(master, ["Instruction", *truth["pipeline"]["cycles"]], expected_pipeline(truth))
        r_pipeline, r_pipeline_evidence = exact_table(review, ["Instruction", *truth["pipeline"]["cycles"]], expected_pipeline(truth))
        m_equation, m_equation_evidence = equation_score(master_entries, truth)
        r_equation, r_equation_evidence = equation_score(review_entries, truth)
        results.append(item("R003", (m_pipeline + r_pipeline + m_equation + r_equation) / 4, f"master[{m_pipeline_evidence}; {m_equation_evidence}]; review[{r_pipeline_evidence}; {r_equation_evidence}]"))
        pair_score, pair_evidence = check_pair(master, review, master_entries, review_entries, truth, master_path, review_path)
        results.append(item("R004", pair_score, pair_evidence))
        master_score, master_control_evidence = check_master(master, truth)
        results.append(item("R005", master_score, master_control_evidence))
        issues_score, issues_evidence = check_review_issues(review, review_entries, truth)
        results.append(item("R006", issues_score, issues_evidence))
        controls_score, controls_evidence = check_controls(master_entries, review_entries, truth)
        results.append(item("R007", controls_score, controls_evidence))
        m_structure, m_structure_evidence = check_structure(master, master_entries, "master", master_render, master_render_evidence)
        r_structure, r_structure_evidence = check_structure(review, review_entries, "review", review_render, review_render_evidence)
        results.append(item("R008", (m_structure + r_structure) / 2, f"master[{m_structure_evidence}]; review[{r_structure_evidence}]"))
        return results, [master_path, review_path]
    except Exception as exc:
        return [item(cid, 0, f"parse error: {type(exc).__name__}: {exc}" if cid == "R001" else "parse error; not scored") for cid in WEIGHTS], [master_path, review_path]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    results, outputs = evaluate(args.artifact_dir)
    values = {entry["criterion_id"]: float(entry["score"]) for entry in results}
    reward = sum(WEIGHTS[cid] * values.get(cid, 0) for cid in WEIGHTS)
    for cid, cap in (("R001", 0.0), ("R002", .32), ("R003", .50), ("R004", .62), ("R005", .72), ("R006", .78), ("R007", .82)):
        if values.get(cid, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    passed = reward >= .86 and all(values.get(cid, 0) == 1 for cid in WEIGHTS)
    payload = {
        "task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(path.read_bytes()).hexdigest() for path in outputs if path.is_file()],
        "criterion_results": results, "reward": reward, "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
