#!/usr/bin/env python3
"""咖啡馆文案审校工作稿的确定性 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from lxml import etree


TASK_ID = "cowork-evo/officebench-1-16-random-paragraph"
OUTPUT_NAME = "cafe_profile_editorial_review.docx"
RUBRIC_VERSION = "officebench-1-16-random-paragraph-v0.3"
JUDGE_VERSION = "officebench-1-16-random-paragraph-deterministic-v0.3"
SOURCE_SHA256 = "4439b9fb95c87edce812407d9961abb9d2928fb1a21668e3c336fb44e9285f55"
TEMPLATE_SHA256 = "f53dfaf5aa724f5ef644a1b9e4ef16f8b7d840c484a0785b3770e3924faa916b"
TICKET_SHA256 = "7f728aacce45dd7884fc3297979ac7c9dab8cc5bf420e1b445bfffad7e6066ed"
WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ0-9]+(?:[’'-][A-Za-zÀ-ÖØ-öø-ÿ0-9]+)*")
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

SPECS = {
    "R001": (8, "hurdle", 0.0),
    "R002": (15, "hurdle", 0.35),
    "R003": (18, "hurdle", 0.58),
    "R004": (10, "hurdle", 0.72),
    "R005": (16, "hurdle", 0.70),
    "R006": (12, "hurdle", 0.76),
    "R007": (7, "positive", None),
    "R008": (7, "positive", None),
    "R009": (5, "positive", None),
    "R010": (2, "positive", None),
}

FIXED_PARAGRAPHS = [
    "Café Profile Editorial Review",
    "Controlled source, approved working copy, and sentence-level evidence",
    "Controlled Source",
    "This block is the fact authority. Preserve it verbatim and use it to check every approved revision.",
    "Approved Working Copy",
    "Apply the approved ticket as five tracked replacements. Keep each sentence ID outside the tracked range.",
    "Sentence Evidence Register",
    "Populate one row per sentence. Word counts use tokens of letters or numbers, allowing internal apostrophes and hyphens.",
    "Copy Integrity Summary",
    "Evidence Note",
]


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def xml_text(element, deleted: bool = False) -> str:
    query = ".//w:delText/text()" if deleted else ".//w:t/text()"
    return "".join(element.xpath(query, namespaces=W_NS))


def split_sentences(text: str) -> list[str]:
    return [item.strip() for item in re.split(r"(?<=[.!?])\s+", text.strip()) if item.strip()]


def words(text: str) -> list[str]:
    return WORD_PATTERN.findall(text)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_directory() -> Path:
    mounted = Path("/app/input_files")
    if (mounted / "random_paragraph.txt").is_file():
        return mounted
    return Path(__file__).resolve().parents[1] / "data/input_files"


def ticket_rows(path: Path) -> list[dict[str, str]]:
    ticket = Document(path)
    table = ticket.tables[0]
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    return [dict(zip(headers, [cell.text.strip() for cell in row.cells])) for row in table.rows[1:]]


def item(identifier: str, fraction: float, evidence: str) -> dict:
    fraction = max(0.0, min(1.0, fraction))
    return {
        "criterion_id": identifier,
        "score": fraction,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {
            "passed": fraction == 1.0,
            "weight": SPECS[identifier][0],
            "type": SPECS[identifier][1],
        },
    }


def valid_docx(path: Path) -> bool:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as package:
            return "word/document.xml" in package.namelist() and "[Content_Types].xml" in package.namelist()
    except Exception:
        return False


def fixed_template_score(template: Document, output: Document, source: str) -> tuple[float, str]:
    paragraphs = [paragraph.text for paragraph in output.paragraphs]
    fixed_hits = sum(paragraphs.count(value) == 1 for value in FIXED_PARAGRAPHS)
    meta_shape = len(output.tables) >= 1 and len(output.tables[0].rows) == 3 and len(output.tables[0].columns) == 2
    meta_fixed = meta_shape and [
        [output.tables[0].cell(row, column).text for column in range(2)] for row in range(2)
    ] == [["Work item", "CP-016"], ["Authority", "random_paragraph.txt"]]
    state_label = meta_shape and output.tables[0].cell(2, 0).text == "Review state"
    source_shape = len(output.tables) >= 2 and len(output.tables[1].rows) == 1 and len(output.tables[1].columns) == 1
    source_exact = source_shape and output.tables[1].cell(0, 0).text.strip() == source
    table_count = len(output.tables) == 4
    shapes = table_count and [(len(table.rows), len(table.columns)) for table in output.tables] == [(3, 2), (1, 1), (6, 6), (4, 2)]
    title_subject = output.core_properties.title == "Café Profile Editorial Review — Approved Copy"
    checks = [fixed_hits / len(FIXED_PARAGRAPHS), meta_fixed, state_label, source_exact, shapes, title_subject]
    return sum(float(value) for value in checks) / len(checks), (
        f"固定段={fixed_hits}/{len(FIXED_PARAGRAPHS)}，元数据={meta_fixed and state_label}，"
        f"受控源={source_exact}，表格形状={shapes}，标题属性={title_subject}"
    )


def tracked_score(output_path: Path, source_sentences: list[str], ticket: list[dict[str, str]]) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
        settings = etree.fromstring(package.read("word/settings.xml"))
    working = root.xpath("//w:body/w:p[w:ins or w:del]", namespaces=W_NS)
    pairs = []
    for index, (source, row) in enumerate(zip(source_sentences, ticket), start=1):
        matches = [
            paragraph
            for paragraph in working
            if "".join(paragraph.xpath("./w:r/w:t/text()", namespaces=W_NS)).strip().startswith(f"S{index} |")
        ]
        if len(matches) != 1:
            pairs.append(False)
            continue
        paragraph = matches[0]
        deleted = "".join(paragraph.xpath("./w:del//w:delText/text()", namespaces=W_NS))
        inserted = "".join(paragraph.xpath("./w:ins//w:t/text()", namespaces=W_NS))
        authors = paragraph.xpath("./w:del/@w:author | ./w:ins/@w:author", namespaces=W_NS)
        revision_shape = len(paragraph.xpath("./w:del", namespaces=W_NS)) == 1 and len(paragraph.xpath("./w:ins", namespaces=W_NS)) >= 1
        pairs.append(
            deleted == source
            and inserted == row["Approved replacement"]
            and revision_shape
            and authors
            and all(author == "Copy Desk" for author in authors)
        )
    exact_counts = len(working) == 5 and len(root.xpath("//w:body/w:p/w:del", namespaces=W_NS)) == 5
    track_enabled = bool(settings.xpath("//w:trackRevisions", namespaces=W_NS))
    correct_order = [
        "".join(paragraph.xpath("./w:r/w:t/text()", namespaces=W_NS)).strip().split(" |")[0]
        for paragraph in working
    ] == [f"S{index}" for index in range(1, 6)]
    fraction = (sum(pairs) + int(exact_counts) + int(track_enabled) + int(correct_order)) / 8
    return fraction, f"修订对={sum(pairs)}/5，精确数量={exact_counts}，Track Changes={track_enabled}，顺序={correct_order}"


def parse_comment_anchors(root) -> dict[str, str]:
    active: list[str] = []
    pieces: dict[str, list[str]] = {}
    for event, node in etree.iterwalk(root, events=("start", "end")):
        if event == "start" and node.tag == qn("w:commentRangeStart"):
            identifier = node.get(qn("w:id"), "")
            active.append(identifier)
            pieces.setdefault(identifier, [])
        elif event == "start" and node.tag == qn("w:t") and active and node.text:
            for identifier in active:
                pieces.setdefault(identifier, []).append(node.text)
        elif event == "end" and node.tag == qn("w:commentRangeEnd"):
            identifier = node.get(qn("w:id"), "")
            if identifier in active:
                active.remove(identifier)
    return {identifier: "".join(value) for identifier, value in pieces.items()}


def comments_score(output_path: Path, ticket: list[dict[str, str]]) -> tuple[float, str]:
    try:
        with zipfile.ZipFile(output_path) as package:
            document = etree.fromstring(package.read("word/document.xml"))
            comments = etree.fromstring(package.read("word/comments.xml"))
    except (KeyError, etree.XMLSyntaxError):
        return 0.0, "缺少可解析的 comments.xml"
    anchors = parse_comment_anchors(document)
    bodies = {
        node.get(qn("w:id"), ""): "".join(node.xpath(".//w:t/text()", namespaces=W_NS))
        for node in comments.xpath("//w:comment", namespaces=W_NS)
    }
    expected = [(row["Comment anchor"], row["Required comment"]) for row in ticket if row["Required comment"]]
    checks = []
    for anchor, body in expected:
        identifiers = [identifier for identifier, value in anchors.items() if value == anchor]
        checks.append(len(identifiers) == 1 and normalize(bodies.get(identifiers[0], "")) == normalize(body))
    exact_count = len(bodies) == 3 and len(anchors) == 3
    return (sum(checks) + int(exact_count)) / 4, f"精确批注={sum(checks)}/3，总数={len(bodies)}，锚定范围={len(anchors)}"


def register_score(output: Document, source_sentences: list[str], ticket: list[dict[str, str]]) -> tuple[float, str]:
    if len(output.tables) < 3:
        return 0.0, "缺少句子证据台账"
    table = output.tables[2]
    expected_headers = ["ID", "Source cue", "Source words", "Approved words", "Delta", "Edit class / review link"]
    if len(table.rows) != 6 or len(table.columns) != 6 or [cell.text for cell in table.rows[0].cells] != expected_headers:
        return 0.0, f"台账形状={len(table.rows)}x{len(table.columns)}"
    passed = 0
    total = 0
    for row, source, ticket_row in zip(table.rows[1:], source_sentences, ticket):
        source_words = words(source)
        approved_words = words(ticket_row["Approved replacement"])
        expected = [
            ticket_row["ID"],
            " ".join(source_words[:6]) + "…",
            str(len(source_words)),
            str(len(approved_words)),
            f"{len(approved_words) - len(source_words):+d}",
        ]
        actual = [cell.text.strip() for cell in row.cells]
        passed += sum(value == wanted for value, wanted in zip(actual[:5], expected))
        total += 5
        passed += int(ticket_row["Edit class"] in actual[5] and f"Review {ticket_row['ID']}" in actual[5])
        total += 1
    return passed / total, f"台账字段={passed}/{total}，形状=6x6"


def integrity_score(output: Document, source_sentences: list[str], ticket: list[dict[str, str]]) -> tuple[float, str]:
    source = " ".join(source_sentences)
    source_total = sum(len(words(value)) for value in source_sentences)
    approved_total = sum(len(words(row["Approved replacement"])) for row in ticket)
    expected = [
        ["Source total", str(source_total)],
        ["Approved total", str(approved_total)],
        ["Net change", f"{approved_total - source_total:+d}"],
        ["Unapproved factual additions", "None"],
    ]
    summary_ok = len(output.tables) >= 4 and [[cell.text.strip() for cell in row.cells] for row in output.tables[3].rows] == expected
    state_ok = len(output.tables) >= 1 and output.tables[0].cell(2, 1).text.strip() == "Approved copy applied"
    allowed = set(FIXED_PARAGRAPHS) | {""} | {f"S{index} | " for index in range(1, 6)}
    unexpected = []
    for paragraph in output.paragraphs:
        text = paragraph.text
        if text in allowed or paragraph.style.name == "Evidence Note" and text.startswith("The five approved revisions"):
            continue
        unexpected.append(text)
    no_extra = not unexpected
    source_unique = sum(source in cell.text for table in output.tables for row in table.rows for cell in row.cells) == 1
    checks = [summary_ok, state_ok, no_extra, source_unique]
    return sum(checks) / len(checks), f"汇总={summary_ok}，状态={state_ok}，无额外段={no_extra}，受控源唯一={source_unique}，额外={unexpected[:2]}"


def navigation_score(output_path: Path, ticket: list[dict[str, str]]) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    passed = 0
    evidence = []
    for row in ticket:
        name = f"review_{row['ID']}"
        bookmarks = root.xpath(f"//w:bookmarkStart[@w:name='{name}']", namespaces=W_NS)
        links = root.xpath(f"//w:hyperlink[@w:anchor='{name}']", namespaces=W_NS)
        paragraphs = bookmarks[0].xpath("ancestor::w:p[1]", namespaces=W_NS) if len(bookmarks) == 1 else []
        content_ok = len(paragraphs) == 1 and row["Approved replacement"] in xml_text(paragraphs[0])
        ok = content_ok and len(links) == 1
        passed += int(ok)
        evidence.append(f"{name}={ok}")
    return passed / 5, ", ".join(evidence)


def geometry(section) -> tuple[int | None, ...]:
    return (
        section.page_width,
        section.page_height,
        section.top_margin,
        section.bottom_margin,
        section.left_margin,
        section.right_margin,
    )


def footer_has_fields(section) -> bool:
    xml = section.footer._element.xml
    return "PAGE" in xml and "NUMPAGES" in xml


def native_score(template: Document, output: Document) -> tuple[float, str]:
    sections_ok = len(output.sections) == 2 and all(
        all(abs((actual or 0) - (expected or 0)) <= 2000 for expected, actual in zip(geometry(t), geometry(o)))
        for t, o in zip(template.sections, output.sections)
    )
    orientations_ok = len(output.sections) == 2 and int(output.sections[0].orientation) == 0 and int(output.sections[1].orientation) == 1
    headers_ok = len(output.sections) == 2 and all(
        normalize(" ".join(p.text for p in expected.header.paragraphs)) == normalize(" ".join(p.text for p in actual.header.paragraphs))
        for expected, actual in zip(template.sections, output.sections)
    )
    footers_ok = len(output.sections) == 2 and all(footer_has_fields(section) for section in output.sections)
    styles = {paragraph.text: paragraph.style.name for paragraph in output.paragraphs if paragraph.text in FIXED_PARAGRAPHS}
    heading_ok = all(styles.get(value) == "Heading 1" for value in ["Controlled Source", "Approved Working Copy", "Sentence Evidence Register", "Copy Integrity Summary", "Evidence Note"])
    working_ok = sum(paragraph.style.name == "Working Sentence" for paragraph in output.paragraphs) == 5
    table = output.tables[2] if len(output.tables) >= 3 else None
    table_native = bool(
        table is not None
        and (
            table.style is not None
            and normalize(table.style.name) == "table grid"
            or table._tbl.xpath("./w:tblPr/w:tblBorders | .//w:tcPr/w:tcBorders")
        )
    )
    repeat = bool(table is not None and table.rows[0]._tr.xpath("./w:trPr/w:tblHeader"))
    nonsplit = bool(table is not None and all(row._tr.xpath("./w:trPr/w:cantSplit") for row in table.rows[1:]))
    checks = [sections_ok, orientations_ok, headers_ok, footers_ok, heading_ok, working_ok, table_native, repeat, nonsplit]
    return sum(checks) / len(checks), (
        f"分节={sections_ok}，方向={orientations_ok}，页眉={headers_ok}，页脚字段={footers_ok}，"
        f"标题={heading_ok}，工作段={working_ok}，原生表={table_native}，重复表头={repeat}，不拆行={nonsplit}"
    )


def evidence_note_score(output: Document) -> tuple[float, str]:
    notes = [paragraph.text for paragraph in output.paragraphs if paragraph.text.startswith("The five approved revisions")]
    if len(notes) != 1:
        return 0.0, f"证据说明数={len(notes)}"
    note = notes[0]
    lower = normalize(note)
    word_count = len(words(note))
    groups = [
        all(term in lower for term in ("s1", "city", "café", "tranquil")),
        all(term in lower for term in ("s2", "vintage furniture", "soft lighting", "locals", "tourists")),
        all(term in lower for term in ("s3", "coffee", "pastry", "aroma")),
        all(term in lower for term in ("s4", "mismatched tables", "conversation", "reflection", "drinks")),
        all(term in lower for term in ("s5", "outside-inside", "time", "respite")),
        all(term in lower for term in ("103", "68", "-35", "proper nouns", "dates", "prices", "menu items", "service promises")),
    ]
    checks = [90 <= word_count <= 130, *groups]
    return sum(checks) / len(checks), f"词数={word_count}，证据组={groups}"


def render_score(path: Path) -> tuple[float, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdftotext") or not shutil.which("pdfinfo"):
        return 0.0, "缺少 LibreOffice/Poppler"
    with tempfile.TemporaryDirectory(prefix="cafe-profile-judge-") as temp:
        root = Path(temp)
        profile = root / "profile"
        command = [
            "libreoffice",
            "--headless",
            f"-env:UserInstallation={profile.as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(root),
            str(path),
        ]
        converted = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False)
        pdf = root / f"{path.stem}.pdf"
        if converted.returncode != 0 or not pdf.is_file():
            return 0.0, f"PDF 转换失败：{converted.stderr[-200:]}"
        text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, timeout=60, check=False).stdout
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=60, check=False).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        required = ["Café Profile Editorial Review", "Sentence Evidence Register", "Approved copy applied", "103", "68", "-35"]
        checks = [pages == 2, len(text.strip()) >= 1000, all(value in text for value in required)]
        return sum(checks) / len(checks), f"页数={pages}，字符={len(text.strip())}，关键渲染={all(value in text for value in required)}"


def safety_score(path: Path) -> tuple[float, str]:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        forbidden = [name for name in names if name.endswith(("vbaProject.bin", ".bin")) or "embeddings/" in name or "activeX/" in name or "altChunk" in name]
        images = [name for name in names if name.startswith("word/media/")]
        external = []
        for name in names:
            if not name.endswith(".rels"):
                continue
            root = etree.fromstring(package.read(name))
            external.extend(node.get("Target", "") for node in root if node.get("TargetMode") == "External")
    safe = not forbidden and not images and not external
    return float(safe), f"禁止部件={forbidden}，图片={images}，外链={external}"


def finalize(results: list[dict], output_path: Path | None) -> tuple[dict, float, bool]:
    scores = {result["criterion_id"]: result["score"] for result in results}
    reward = sum(SPECS[identifier][0] * scores.get(identifier, 0.0) for identifier in SPECS) / 100
    caps = [spec[2] for identifier, spec in SPECS.items() if spec[1] == "hurdle" and scores.get(identifier, 0.0) < 1.0 and spec[2] is not None]
    if caps:
        reward = min(reward, min(caps))
    reward = round(reward, 6)
    passed = reward >= 0.8 and all(scores.get(identifier, 0.0) == 1.0 for identifier in ("R001", "R002", "R003", "R004", "R005", "R006"))
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [sha256(output_path)] if output_path and output_path.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": passed,
    }
    return payload, reward, passed


def evaluate(artifact_dir: Path) -> tuple[dict, float, bool]:
    source_dir = source_directory()
    output_path = artifact_dir / OUTPUT_NAME
    visible = sorted(path.name for path in artifact_dir.iterdir() if path.is_file()) if artifact_dir.is_dir() else []
    valid = valid_docx(output_path)
    exact_delivery = visible == [OUTPUT_NAME]
    results = [item("R001", float(valid and exact_delivery), f"文件={visible}，原生DOCX={valid}")]
    if not valid:
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, "核心 DOCX 缺失或无法解析"))
        return finalize(results, output_path if output_path.is_file() else None)

    source_path = source_dir / "random_paragraph.txt"
    template_path = source_dir / "cafe_profile_editorial_template.docx"
    ticket_path = source_dir / "cafe_profile_approved_copy_ticket.docx"
    input_hashes = (sha256(source_path), sha256(template_path), sha256(ticket_path))
    if input_hashes != (SOURCE_SHA256, TEMPLATE_SHA256, TICKET_SHA256):
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, f"冻结输入校验失败：{input_hashes}"))
        return finalize(results, output_path)

    source = source_path.read_text(encoding="utf-8").strip()
    source_sentences = split_sentences(source)
    ticket = ticket_rows(ticket_path)
    template = Document(template_path)
    output = Document(output_path)

    fraction, evidence = fixed_template_score(template, output, source)
    results.append(item("R002", fraction, evidence))
    fraction, evidence = tracked_score(output_path, source_sentences, ticket)
    results.append(item("R003", fraction, evidence))
    fraction, evidence = comments_score(output_path, ticket)
    results.append(item("R004", fraction, evidence))
    fraction, evidence = register_score(output, source_sentences, ticket)
    results.append(item("R005", fraction, evidence))
    fraction, evidence = integrity_score(output, source_sentences, ticket)
    results.append(item("R006", fraction, evidence))
    fraction, evidence = navigation_score(output_path, ticket)
    results.append(item("R007", fraction, evidence))
    fraction, evidence = native_score(template, output)
    results.append(item("R008", fraction, evidence))
    fraction, evidence = evidence_note_score(output)
    results.append(item("R009", fraction, evidence))
    render_fraction, render_evidence = render_score(output_path)
    safe_fraction, safe_evidence = safety_score(output_path)
    results.append(item("R010", (render_fraction + safe_fraction) / 2, f"{render_evidence}；{safe_evidence}"))
    return finalize(results, output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    try:
        payload, reward, passed = evaluate(args.artifact_dir)
    except Exception as exc:
        results = [item("R001", 0.0, f"Judge 读取失败：{type(exc).__name__}: {exc}")]
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, "未评分"))
        payload, reward, passed = finalize(results, None)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
