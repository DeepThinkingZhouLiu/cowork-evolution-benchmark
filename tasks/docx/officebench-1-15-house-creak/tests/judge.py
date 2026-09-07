#!/usr/bin/env python3
"""The House that Creaked 编辑工作稿的确定性 Judge。"""

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


TASK_ID = "cowork-evo/officebench-1-15-house-creak"
OUTPUT_NAME = "house_creak_editorial_review.docx"
RUBRIC_VERSION = "officebench-1-15-house-creak-v0.3"
JUDGE_VERSION = "officebench-1-15-house-creak-deterministic-v0.3"
SOURCE_SHA256 = "e549d600d201cefcf67cc4e42cf04ff1b2828d567689cb36a29af8e6e7655fd7"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

SPECS = {
    "R001": (12, "hurdle", 0.0),
    "R002": (20, "hurdle", 0.35),
    "R003": (9, "positive", None),
    "R004": (10, "positive", None),
    "R005": (14, "hurdle", 0.72),
    "R006": (11, "hurdle", 0.72),
    "R007": (10, "hurdle", 0.72),
    "R008": (6, "positive", None),
    "R009": (4, "positive", None),
    "R010": (4, "positive", None),
}

BOOKMARKS = [
    "move_victorian_gothic",
    "creaking_bedrooms",
    "electrical_sparking",
    "bed_ceiling_collapse",
    "gas_explosion",
    "condemned_disclosure",
]

COMMENT_REQUIREMENTS = {
    "power point which didn't spark whenever an appliance requiring over 1000 watts was plugged into it.": ("electrical", "warning"),
    "a double bed crash through the dining room ceiling.": ("structural", "failure"),
    "That house was condemned years ago!": ("risk", "confirm"),
}

CHRONOLOGY_REQUIREMENTS = [
    ("victorian gothic", ("remote", "property"), ("unsafe", "isolated")),
    ("sound was worst in the master bedroom", ("movement", "room"), ("warning", "sign")),
    ("over 1000 watts", ("electrical", "spark"), ("hazard", "catastrophe")),
    ("double bed crash through the dining room ceiling", ("structural", "failure"), ("ros", "room")),
    ("boom", ("gas", "oven"), ("walls", "fire")),
    ("house was condemned years ago", ("pre-existing", "condition"), ("confirm", "preventable")),
]

CHARACTER_REQUIREMENTS = {
    "ros": (("bark", "creak"), ("mother", "dining"), ("warning", "human")),
    "mother": (("property", "outside"), ("children", "decision"), ("agency", "acquiescence")),
    "peter and amanda": (("lesson", "mother"), ("grandmother", "school"), ("depend", "risk")),
    "father": (("aspiration", "concern"), ("uninsured", "choice"), ("authority", "evidence")),
    "neighbour and women's institute": (("condemn", "food"), ("emergency", "support"), ("knowledge", "due diligence")),
}


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def result_item(identifier: str, passed: bool, evidence: str, fraction: float = 1.0) -> dict:
    score = max(0.0, min(1.0, fraction if passed else 0.0))
    return {
        "criterion_id": identifier,
        "score": score,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"passed": bool(passed), "weight": SPECS[identifier][0], "type": SPECS[identifier][1]},
    }


def run_format_signature(paragraph) -> list[tuple]:
    merged = []
    for run in paragraph.runs:
        if not run.text:
            continue
        signature = (
            run.bold,
            run.italic,
            run.underline,
            run.font.name,
            run.font.size.pt if run.font.size else None,
            str(run.font.color.rgb) if run.font.color and run.font.color.rgb else None,
        )
        if merged and merged[-1][0] == signature:
            merged[-1] = (signature, merged[-1][1] + run.text)
        else:
            merged.append((signature, run.text))
    return merged


def paragraph_signature(paragraph) -> tuple:
    fmt = paragraph.paragraph_format
    return (
        paragraph.style.name if paragraph.style else None,
        paragraph.alignment,
        fmt.left_indent,
        fmt.right_indent,
        fmt.first_line_indent,
        fmt.space_before,
        fmt.space_after,
        fmt.line_spacing,
        tuple(run_format_signature(paragraph)),
    )


def source_dir() -> Path:
    container = Path("/app/input_files")
    if (container / "House_creak.docx").is_file():
        return container
    return Path(__file__).resolve().parents[1] / "data/input_files"


def section_geometry(section) -> tuple:
    return (
        section.page_width,
        section.page_height,
        section.top_margin,
        section.right_margin,
        section.bottom_margin,
        section.left_margin,
    )


def footer_has_page_field(section) -> bool:
    xml = section.footer._element.xml
    return "PAGE" in xml or any(paragraph.text.strip().isdigit() for paragraph in section.footer.paragraphs)


def heading_following_paragraph(document, heading: str) -> str:
    for index, paragraph in enumerate(document.paragraphs):
        if paragraph.text.strip() == heading:
            for following in document.paragraphs[index + 1 :]:
                if following.text.strip():
                    return following.text.strip()
    return ""


def find_table(document, expected_headers: list[str]):
    wanted = [normalize(value) for value in expected_headers]
    for table in document.tables:
        if not table.rows:
            continue
        actual = [normalize(cell.text) for cell in table.rows[0].cells]
        if actual == wanted:
            return table
    return None


def cells_contain(text: str, tokens: tuple[str, ...]) -> bool:
    value = normalize(text)
    return all(token.lower() in value for token in tokens)


def parse_comment_anchors(document_xml: bytes) -> dict[str, str]:
    root = etree.fromstring(document_xml)
    anchors: dict[str, str] = {}
    for paragraph in root.xpath("//w:body//w:p", namespaces=W_NS):
        active: list[str] = []
        pieces: dict[str, list[str]] = {}
        for node in paragraph.iter():
            if node.tag == qn("w:commentRangeStart"):
                identifier = node.get(qn("w:id"), "")
                active.append(identifier)
                pieces.setdefault(identifier, [])
            elif node.tag == qn("w:t"):
                for identifier in active:
                    pieces.setdefault(identifier, []).append(node.text or "")
            elif node.tag == qn("w:commentRangeEnd"):
                identifier = node.get(qn("w:id"), "")
                if identifier in active:
                    active.remove(identifier)
        for identifier, values in pieces.items():
            anchors[identifier] = "".join(values)
    return anchors


def parse_comments(comments_xml: bytes) -> dict[str, str]:
    root = etree.fromstring(comments_xml)
    return {
        node.get(qn("w:id"), ""): "".join(node.xpath(".//w:t/text()", namespaces=W_NS))
        for node in root.xpath("//w:comment", namespaces=W_NS)
    }


def render_check(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    pdfinfo = shutil.which("pdfinfo")
    if not all((libreoffice, pdftotext, pdfinfo)):
        return False, "缺少 LibreOffice/Poppler 渲染依赖"
    with tempfile.TemporaryDirectory(prefix="house-creak-judge-") as temp:
        root = Path(temp)
        profile = (root / "profile").resolve().as_uri()
        converted = subprocess.run(
            [
                libreoffice,
                "--headless",
                f"-env:UserInstallation={profile}",
                "--convert-to",
                "pdf",
                "--outdir",
                temp,
                str(path),
            ],
            capture_output=True,
            text=True,
            timeout=240,
            check=False,
        )
        pdf = root / f"{path.stem}.pdf"
        if converted.returncode or not pdf.is_file():
            return False, f"PDF 转换失败：{(converted.stderr or converted.stdout)[-500:]}"
        text = subprocess.run([pdftotext, "-layout", str(pdf), "-"], capture_output=True, text=True, timeout=120, check=False)
        info = subprocess.run([pdfinfo, str(pdf)], capture_output=True, text=True, timeout=60, check=False)
        if text.returncode or info.returncode:
            return False, "PDF 文本或页数检查失败"
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        page_texts = text.stdout.split("\f")
        substantive = [len(re.sub(r"\s+", "", item)) for item in page_texts if item.strip()]
        required = ["The House that Creaked", "Editorial Analysis Appendix", "Incident Chronology", "Character Decision Matrix"]
        missing = [item for item in required if item not in text.stdout]
        near_blank = [index + 1 for index, count in enumerate(substantive) if count < 80]
        ok = 6 <= pages <= 12 and len(substantive) == pages and not near_blank and not missing
        return ok, f"页数={pages}，逐页字符={substantive}，近空白页={near_blank}，缺失标题={missing}"


def evaluate(artifact_dir: Path) -> dict:
    output = artifact_dir / OUTPUT_NAME
    items: list[dict] = []
    if not output.is_file() or not zipfile.is_zipfile(output):
        items.append(result_item("R001", False, f"缺少或损坏 {OUTPUT_NAME}"))
        return finalize(items, output)

    source_path = source_dir() / "House_creak.docx"
    input_digest = hashlib.sha256(source_path.read_bytes()).hexdigest()
    try:
        source = Document(source_path)
        candidate = Document(output)
        with zipfile.ZipFile(output) as archive:
            names = archive.namelist()
            document_xml = archive.read("word/document.xml")
            comments_xml = archive.read("word/comments.xml") if "word/comments.xml" in names else b""
            rel_payloads = [archive.read(name) for name in names if name.endswith(".rels")]
    except Exception as error:
        items.append(result_item("R001", False, f"DOCX 解析失败：{error}"))
        return finalize(items, output)

    editable = bool(candidate.paragraphs) and len(candidate.tables) >= 2
    items.append(result_item("R001", editable, f"输入SHA={input_digest}，段落={len(candidate.paragraphs)}，表格={len(candidate.tables)}"))

    source_text = [paragraph.text for paragraph in source.paragraphs]
    candidate_text = [paragraph.text for paragraph in candidate.paragraphs[: len(source_text)]]
    text_preserved = candidate_text == source_text and len(source_text) == 47
    source_formats = [paragraph_signature(paragraph) for paragraph in source.paragraphs]
    candidate_formats = [paragraph_signature(paragraph) for paragraph in candidate.paragraphs[: len(source_text)]]
    format_preserved = source_formats == candidate_formats
    geometry_preserved = bool(candidate.sections) and all(
        abs(left - right) <= 2000
        for left, right in zip(section_geometry(candidate.sections[0]), section_geometry(source.sections[0]))
    )
    page_numbers_preserved = bool(candidate.sections) and footer_has_page_field(candidate.sections[0]) and footer_has_page_field(source.sections[0])
    preserved = input_digest == SOURCE_SHA256 and text_preserved and format_preserved and geometry_preserved and page_numbers_preserved
    items.append(result_item("R002", preserved, f"原段落={len(source_text)}，文字={text_preserved}，格式={format_preserved}，页面几何={geometry_preserved}，页码={page_numbers_preserved}"))

    body_root = etree.fromstring(document_xml)
    toc_instructions = " ".join(body_root.xpath("//w:instrText/text() | //w:fldSimple/@w:instr", namespaces=W_NS))
    heading_styles = {
        paragraph.text.strip(): paragraph.style.name if paragraph.style else ""
        for paragraph in candidate.paragraphs
        if paragraph.text.strip() in {"Editorial Analysis Appendix", "Synopsis", "Incident Chronology", "Character Decision Matrix", "Editorial Note"}
    }
    section_ok = len(candidate.sections) == 2 and str(candidate.sections[1].start_type).endswith("NEW_PAGE (2)")
    header_text = " ".join(paragraph.text for paragraph in candidate.sections[-1].header.paragraphs)
    headings_ok = heading_styles.get("Editorial Analysis Appendix") == "Heading 1" and all(
        heading_styles.get(name) == "Heading 2" for name in ("Synopsis", "Incident Chronology", "Character Decision Matrix", "Editorial Note")
    )
    structure_ok = section_ok and "TOC" in toc_instructions and '1-2' in toc_instructions and headings_ok and header_text.strip() == "Anthology Editorial Review — The House that Creaked"
    items.append(result_item("R003", structure_ok, f"分节={len(candidate.sections)} / {section_ok}，TOC={toc_instructions!r}，标题={heading_styles}，页眉={header_text!r}"))

    synopsis = heading_following_paragraph(candidate, "Synopsis")
    word_count = len(re.findall(r"\b[A-Za-z]+(?:['’-][A-Za-z]+)*\b", synopsis))
    synopsis_groups = [
        ("move", "victorian"),
        ("ros", "warning"),
        ("creak", "1000"),
        ("bed", "ceiling", "explosion"),
        ("condemn", "neighbour"),
        ("grandmother", "family"),
    ]
    synopsis_norm = normalize(synopsis)
    group_checks = [all(token in synopsis_norm for token in group) for group in synopsis_groups]
    contradictions = [phrase for phrase in ("father warned the family", "house was insured", "ros died", "family returned to the mansion") if phrase in synopsis_norm]
    synopsis_ok = 120 <= word_count <= 180 and all(group_checks) and not contradictions
    items.append(result_item("R004", synopsis_ok, f"词数={word_count}，因果链={group_checks}，冲突断言={contradictions}"))

    chronology = find_table(candidate, ["Seq", "Story evidence", "Risk signal", "Consequence", "Source link"])
    chronology_checks = []
    if chronology is not None and len(chronology.rows) == 7:
        for index, requirement in enumerate(CHRONOLOGY_REQUIREMENTS, start=1):
            row = chronology.rows[index].cells
            evidence_token, risk_tokens, consequence_tokens = requirement
            chronology_checks.append(
                normalize(row[0].text) == str(index)
                and evidence_token in normalize(row[1].text)
                and cells_contain(row[2].text, risk_tokens)
                and cells_contain(row[3].text, consequence_tokens)
            )
    chronology_ok = len(chronology_checks) == 6 and all(chronology_checks)
    items.append(result_item("R005", chronology_ok, f"表存在={chronology is not None}，行数={len(chronology.rows) if chronology else 0}，逐行={chronology_checks}"))

    characters = find_table(candidate, ["Character", "Evidence-based action or belief", "Effect on events", "Editorial interpretation"])
    character_checks = {}
    if characters is not None:
        observed = {normalize(row.cells[0].text): row.cells for row in characters.rows[1:]}
        for role, groups in CHARACTER_REQUIREMENTS.items():
            cells = observed.get(role, [])
            character_checks[role] = bool(cells) and all(cells_contain(cells[column + 1].text, tokens) for column, tokens in enumerate(groups))
    characters_ok = len(character_checks) == 5 and all(character_checks.values()) and len(characters.rows) == 6
    items.append(result_item("R006", characters_ok, f"表存在={characters is not None}，逐角色={character_checks}"))

    anchors = parse_comment_anchors(document_xml)
    comments = parse_comments(comments_xml) if comments_xml else {}
    comment_checks = {}
    for phrase, tokens in COMMENT_REQUIREMENTS.items():
        matching = [identifier for identifier, anchor in anchors.items() if anchor == phrase]
        comment_checks[phrase] = len(matching) == 1 and matching[0] in comments and cells_contain(comments[matching[0]], tokens)
    comments_ok = all(comment_checks.values()) and len(comments) == 3
    items.append(result_item("R007", comments_ok, f"批注数={len(comments)}，锚定检查={comment_checks}"))

    bookmark_names = set(body_root.xpath("//w:bookmarkStart/@w:name", namespaces=W_NS))
    link_anchors = body_root.xpath("//w:hyperlink/@w:anchor", namespaces=W_NS)
    links_ok = all(name in bookmark_names for name in BOOKMARKS) and sorted(link_anchors) == sorted(BOOKMARKS)
    items.append(result_item("R008", links_ok, f"书签={sorted(bookmark_names)}，内链锚点={sorted(link_anchors)}"))

    dangerous_names = [name for name in names if "vbaproject" in name.lower() or name.startswith("word/embeddings/")]
    external_rels = sum(payload.count(b'TargetMode="External"') for payload in rel_payloads)
    media = [name for name in names if name.startswith("word/media/") and not name.endswith("/")]
    table_nodes = body_root.xpath("//w:body/w:tbl", namespaces=W_NS)
    table_object_checks = []
    for table_node in table_nodes:
        rows = table_node.xpath("./w:tr", namespaces=W_NS)
        repeat_header = bool(rows and rows[0].xpath("./w:trPr/w:tblHeader", namespaces=W_NS))
        no_split = all(row.xpath("./w:trPr/w:cantSplit", namespaces=W_NS) for row in rows[1:])
        cells = table_node.xpath(".//w:tc", namespaces=W_NS)
        grid = bool(table_node.xpath("./w:tblPr/w:tblBorders", namespaces=W_NS)) or (
            bool(cells) and len(table_node.xpath(".//w:tcPr/w:tcBorders", namespaces=W_NS)) == len(cells)
        )
        table_object_checks.append(repeat_header and no_split and grid)
    safe_editable = not dangerous_names and external_rels == 0 and not media and len(table_nodes) == 2 and all(table_object_checks)
    items.append(result_item("R009", safe_editable, f"危险部件={dangerous_names}，外链={external_rels}，媒体={media}，表对象={table_object_checks}"))

    render_ok, render_evidence = render_check(output)
    items.append(result_item("R010", render_ok, render_evidence))
    return finalize(items, output)


def finalize(items: list[dict], output: Path) -> dict:
    present = {item["criterion_id"] for item in items}
    for identifier in SPECS:
        if identifier not in present:
            items.append(result_item(identifier, False, "因为文件缺失、损坏或前置门槛失败，未执行该项"))
    items.sort(key=lambda item: item["criterion_id"])
    total = sum(spec[0] for spec in SPECS.values())
    earned = sum(SPECS[item["criterion_id"]][0] * item["score"] for item in items)
    reward = earned / total
    hurdle_failures = [item["criterion_id"] for item in items if SPECS[item["criterion_id"]][1] == "hurdle" and item["score"] < 1]
    if hurdle_failures:
        reward = min(reward, min(SPECS[identifier][2] for identifier in hurdle_failures if SPECS[identifier][2] is not None))
    reward = round(max(0.0, min(1.0, reward)), 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": items,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    payload = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.write_text(f"{payload['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
