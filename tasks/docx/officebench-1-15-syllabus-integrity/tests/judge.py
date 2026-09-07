#!/usr/bin/env python3
"""课程政策整合红线工作稿的确定性 Judge。"""

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
from docx.text.paragraph import Paragraph
from lxml import etree


TASK_ID = "cowork-evo/officebench-1-15-syllabus-integrity"
OUTPUT_NAME = "sample_syllabus_policy_review.docx"
RUBRIC_VERSION = "officebench-1-15-syllabus-integrity-v0.3"
JUDGE_VERSION = "officebench-1-15-syllabus-integrity-deterministic-v0.3"
MASTER_SHA256 = "9efefdcac29942e59edf9537aa5cfade576ea55b408dde34205792fd773c8505"
POLICY_SHA256 = "260bd988b35c3a3577238d9f15f6a15010a39e5de3a167cc73b65bbaad3291cc"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
R_NS = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

SPECS = {
    "R001": (10, "hurdle", 0.0),
    "R002": (18, "hurdle", 0.35),
    "R003": (15, "hurdle", 0.62),
    "R004": (8, "positive", None),
    "R005": (10, "hurdle", 0.78),
    "R006": (15, "hurdle", 0.72),
    "R007": (9, "hurdle", 0.78),
    "R008": (6, "positive", None),
    "R009": (6, "positive", None),
    "R010": (3, "positive", None),
}

BOOKMARKS = [
    "general_research_misconduct",
    "research_data_scope",
    "research_study_purpose",
    "research_participant_selection",
    "research_risk_benefit",
    "research_anonymity",
    "off_campus_application",
]

COMMENT_CHECKS = {
    "falsification or sabotage of research data": ("general", "misconduct", "research"),
    "verified using Turnitin technology": ("originality", "verification", "research"),
    "apply to your interactions with the staff of the schools as well as the community organizations and its clients.": ("field", "community", "research"),
}

MATRIX_CHECKS = [
    (
        "general research-data misconduct",
        ("fabrication", "falsification", "sabotage", "research data"),
        ("misconduct", "sanction"),
        "general_research_misconduct",
    ),
    (
        "research-data scope",
        ("students", "teachers", "classmates", "orally", "writing", "observation", "records", "artifacts", "research data"),
        ("professionally", "ethically"),
        "research_data_scope",
    ),
    (
        "study purpose",
        ("clear", "questions", "data"),
        ("define", "explain", "answer"),
        "research_study_purpose",
    ),
    (
        "unbiased participant selection",
        ("inclusion", "exclusion", "unbiased", "extraneous", "unrelated"),
        ("selection", "research questions"),
        "research_participant_selection",
    ),
    (
        "risk/benefit balance",
        ("risks", "discomforts", "hazards", "benefits", "balanced"),
        ("state", "evaluate", "balance"),
        "research_risk_benefit",
    ),
    (
        "anonymity",
        ("all participants", "all times"),
        ("protect", "anonymity"),
        "research_anonymity",
    ),
    (
        "field/community application",
        ("university policy", "code", "staff", "community organizations", "clients"),
        ("field", "clinical", "internship", "community"),
        "off_campus_application",
    ),
]


def normalized(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().lower()


def xml_text(element) -> str:
    try:
        values = element.xpath(".//w:t/text()", namespaces=W_NS)
    except TypeError:
        values = element.xpath(".//w:t/text()")
    return "".join(values)


def criterion(identifier: str, fraction: float, evidence: str) -> dict:
    score = max(0.0, min(1.0, fraction))
    return {
        "criterion_id": identifier,
        "score": score,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {
            "passed": score == 1.0,
            "weight": SPECS[identifier][0],
            "type": SPECS[identifier][1],
        },
    }


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def valid_docx(path: Path) -> bool:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as package:
            return "word/document.xml" in package.namelist() and "[Content_Types].xml" in package.namelist()
    except Exception:
        return False


def source_directory() -> Path:
    container = Path("/app/input_files")
    if (container / "sample_syllabus.docx").is_file():
        return container
    return Path(__file__).resolve().parents[1] / "data/input_files"


def direct_body_paragraphs(document) -> list:
    return document.element.body.xpath("./w:p")


def style_name(document, paragraph_element) -> str:
    return Paragraph(paragraph_element, document._body).style.name


def run_signature(paragraph_element) -> list[tuple]:
    def canonical_run_style(value: str | None) -> str | None:
        if value is None:
            return None
        compact = re.sub(r"[^a-z0-9]", "", value.lower())
        # LibreOffice 对两个既有字符样式只改名，不改变其链接/空格语义。
        if compact == "internetlink":
            return "hyperlink"
        return compact

    merged: list[tuple] = []
    for run in paragraph_element.xpath(".//w:r[not(w:commentReference)]"):
        text = "".join(run.xpath(".//w:t/text()"))
        if not text:
            continue
        rpr = run.find(qn("w:rPr"))
        signature = (
            bool(rpr is not None and rpr.find(qn("w:b")) is not None),
            bool(rpr is not None and rpr.find(qn("w:i")) is not None),
            bool(rpr is not None and rpr.find(qn("w:u")) is not None),
            (rpr.find(qn("w:sz")).get(qn("w:val")) if rpr is not None and rpr.find(qn("w:sz")) is not None else None),
            canonical_run_style(rpr.find(qn("w:rStyle")).get(qn("w:val")) if rpr is not None and rpr.find(qn("w:rStyle")) is not None else None),
        )
        if merged and merged[-1][0] == signature:
            merged[-1] = (signature, merged[-1][1] + text)
        else:
            merged.append((signature, text))
    return merged


def media_hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as package:
        return {
            name: hashlib.sha256(package.read(name)).hexdigest()
            for name in package.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        }


def external_targets(path: Path) -> list[str]:
    values: list[str] = []
    with zipfile.ZipFile(path) as package:
        for name in package.namelist():
            if not name.endswith(".rels"):
                continue
            root = etree.fromstring(package.read(name))
            for relationship in root:
                if relationship.get("TargetMode") == "External":
                    values.append(relationship.get("Target", ""))
    return sorted(values)


def geometry(section) -> tuple:
    return (
        section.page_width,
        section.page_height,
        section.top_margin,
        section.bottom_margin,
        section.left_margin,
        section.right_margin,
        section.header_distance,
        section.footer_distance,
    )


def protect_master(master_path: Path, output_path: Path, master, output) -> tuple[float, str]:
    source_paragraphs = direct_body_paragraphs(master)
    output_paragraphs = direct_body_paragraphs(output)
    appendix_index = next((i for i, p in enumerate(output_paragraphs) if xml_text(p) == "Policy Integration Review"), -1)
    before_appendix = output_paragraphs[:appendix_index] if appendix_index >= 0 else output_paragraphs
    protected_candidates = []
    for paragraph in before_appendix:
        if paragraph.xpath("./w:ins[normalize-space(@w:author)='Course Policy Editor']"):
            continue
        protected_candidates.append(paragraph)

    # python-docx 保存时会为新节增加独立空段；LibreOffice 另存时会把同一个
    # sectPr 合并到主模板末尾原本就存在的空段。只在确有多出的第 188 段时
    # 去掉该空分节段，避免把引擎的等价序列化误判为模板删段。
    if len(protected_candidates) == len(source_paragraphs) + 1:
        extra_sections = [
            paragraph
            for paragraph in protected_candidates
            if not xml_text(paragraph) and paragraph.xpath("./w:pPr/w:sectPr")
        ]
        if len(extra_sections) == 1:
            protected_candidates.remove(extra_sections[0])

    text_ok = len(protected_candidates) == len(source_paragraphs) and all(
        xml_text(actual) == xml_text(expected)
        for expected, actual in zip(source_paragraphs, protected_candidates)
    )
    style_ok = len(protected_candidates) == len(source_paragraphs) and all(
        style_name(master, expected) == style_name(output, actual)
        for expected, actual in zip(source_paragraphs, protected_candidates)
    )
    format_matches = sum(
        run_signature(expected) == run_signature(actual)
        for expected, actual in zip(source_paragraphs, protected_candidates)
    )
    format_fraction = format_matches / len(source_paragraphs) if source_paragraphs else 0.0

    schedule_ok = False
    if len(master.tables) == 1 and output.tables:
        expected = [[cell.text for cell in row.cells] for row in master.tables[0].rows]
        actual = [[cell.text for cell in row.cells] for row in output.tables[0].rows]
        schedule_ok = len(output.tables[0].rows) == 34 and len(output.tables[0].columns) == 4 and expected == actual
    media_ok = media_hashes(master_path) == media_hashes(output_path)
    external_ok = external_targets(master_path) == external_targets(output_path)
    core_geometry_ok = len(output.sections) >= 2 and all(
        abs((actual or 0) - (expected or 0)) <= 2000
        for expected, actual in zip(geometry(master.sections[0])[:6], geometry(output.sections[0])[:6])
    )
    source_header = "|".join(paragraph.text for paragraph in master.sections[0].header.paragraphs)
    source_footer = "|".join(paragraph.text for paragraph in master.sections[0].footer.paragraphs)
    output_header = "|".join(paragraph.text for paragraph in output.sections[0].header.paragraphs)
    output_footer = "|".join(paragraph.text for paragraph in output.sections[0].footer.paragraphs)
    geometry_ok = core_geometry_ok and source_header == output_header and source_footer == output_footer
    checks = [text_ok, style_ok, format_fraction >= 0.98, schedule_ok, media_ok, external_ok, geometry_ok]
    return sum(checks) / len(checks), (
        f"保护段={len(protected_candidates)}/{len(source_paragraphs)}，文字={text_ok}，样式={style_ok}，run格式={format_matches}/{len(source_paragraphs)}，"
        f"课程表={schedule_ok}，媒体={media_ok}，外链={external_ok}，页面几何={geometry_ok}"
    )


def tracked_policy(policy, output) -> tuple[float, float, str]:
    expected_paragraphs = [p.text for p in policy.paragraphs if p.text.strip()]
    output_paragraphs = direct_body_paragraphs(output)
    texts = [xml_text(p) for p in output_paragraphs]
    turnitin = next((i for i, text in enumerate(texts) if "verified using Turnitin technology" in text), -1)
    special = next((i for i, text in enumerate(texts) if text.startswith("SPECIAL NEEDS Accommodations")), -1)
    found_indexes: list[int] = []
    cursor = turnitin + 1
    for expected in expected_paragraphs:
        try:
            found = texts.index(expected, cursor)
        except ValueError:
            found = -1
        found_indexes.append(found)
        cursor = found + 1 if found >= 0 else cursor
    exact_sequence = all(index >= 0 for index in found_indexes)
    correct_anchor = exact_sequence and turnitin >= 0 and special >= 0 and found_indexes == list(range(turnitin + 1, turnitin + 1 + len(expected_paragraphs))) and found_indexes[-1] < special
    tracked_flags = []
    for expected, index in zip(expected_paragraphs, found_indexes):
        if index < 0:
            tracked_flags.append(False)
            continue
        inserted_texts = output_paragraphs[index].xpath("./w:ins//w:t/text()")
        tracked_flags.append("".join(inserted_texts) == expected)
    track_enabled = bool(output.settings.element.xpath("./w:trackRevisions"))
    duplicates_ok = all(texts.count(expected) == 1 for expected in expected_paragraphs)
    tracked_fraction = sum(tracked_flags) / len(expected_paragraphs) if expected_paragraphs else 0.0
    r3_checks = [exact_sequence, correct_anchor, tracked_fraction == 1.0, track_enabled, duplicates_ok]

    style_ok = False
    list_ok = False
    if exact_sequence:
        styles = [style_name(output, output_paragraphs[index]) for index in found_indexes]
        style_ok = styles[0] == "Heading 2" and all(name == "Body Text" for name in styles[1:3]) and all(name == "List Paragraph" for name in styles[3:])
        list_ok = all(
            bool(output_paragraphs[index].xpath("./w:pPr/w:numPr/w:numId"))
            for index in found_indexes[3:]
        )
    r4_checks = [exact_sequence, style_ok, list_ok]
    evidence = (
        f"Turnitin={turnitin}，SpecialNeeds={special}，插入位置={found_indexes}，"
        f"跟踪段={sum(tracked_flags)}/{len(expected_paragraphs)}，trackRevisions={track_enabled}，唯一={duplicates_ok}，样式={style_ok}，编号={list_ok}"
    )
    return sum(r3_checks) / len(r3_checks), sum(r4_checks) / len(r4_checks), evidence


def scope_decision(output) -> tuple[float, str]:
    paragraphs = direct_body_paragraphs(output)
    texts = [xml_text(p) for p in paragraphs]
    try:
        index = texts.index("Scope Decision")
        scope = texts[index + 1]
    except (ValueError, IndexError):
        return 0.0, "未找到 Scope Decision 正文"
    lower = normalized(scope)
    words = re.findall(r"\b[\w’'-]+\b", scope)
    groups = [
        all(term in lower for term in ("general academic integrity", "falsification", "sabotage", "research data", "sanctions")),
        all(term in lower for term in ("turnitin", "originality", "written work", "pledge", "consent")),
        all(term in lower for term in ("students", "teachers", "classmates", "oral", "written", "observ", "records", "artifacts", "unbiased", "risk-benefit", "anonymity")),
        all(term in lower for term in ("off-campus", "academic integrity policy", "code of student responsibility", "staff", "community organizations", "clients")),
        all(term in lower for term in ("cumulatively", "does not substitute", "does not narrow")),
        all(term in lower for term in ("template", "not a representation", "placeholders")),
    ]
    forbidden = re.findall(
        r"\b(?:irb|institutional review board|informed consent|participant consent|consent form|data retention|retention period|mandatory training|research approval)\b",
        lower,
    )
    checks = [120 <= len(words) <= 180, all(groups), not forbidden]
    return sum(checks) / len(checks), f"词数={len(words)}，事实组={groups}，无来源规则={forbidden}"


def find_matrix(output):
    for table in output.tables:
        if len(table.rows) < 1 or len(table.columns) != 4:
            continue
        header = [normalized(cell.text) for cell in table.rows[0].cells]
        if header == ["control", "evidence-based scope", "operational consequence", "source link"]:
            return table
    return None


def matrix_score(output) -> tuple[float, dict[str, bool]]:
    table = find_matrix(output)
    if table is None:
        return 0.0, {name: False for name, *_ in MATRIX_CHECKS}
    rows = {normalized(row.cells[0].text): row for row in table.rows[1:]}
    results: dict[str, bool] = {}
    for name, scope_terms, consequence_terms, anchor in MATRIX_CHECKS:
        row = rows.get(name)
        if row is None:
            results[name] = False
            continue
        scope = normalized(row.cells[1].text)
        consequence = normalized(row.cells[2].text)
        results[name] = all(term in scope for term in scope_terms) and all(term in consequence for term in consequence_terms)
    exact_shape = len(table.rows) == 8 and len(table.columns) == 4
    return (sum(results.values()) + int(exact_shape)) / 8, results


def comments_score(output_path: Path) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        if "word/comments.xml" not in package.namelist():
            return 0.0, "缺 comments.xml"
        comments_root = etree.fromstring(package.read("word/comments.xml"))
        document_root = etree.fromstring(package.read("word/document.xml"))
    comment_text = {
        node.get(qn("w:id")): normalized(xml_text(node))
        for node in comments_root.xpath("//w:comment", namespaces=W_NS)
    }
    anchors: dict[str, str] = {}
    active: list[str] = []
    fragments: dict[str, list[str]] = {}
    for event, node in etree.iterwalk(document_root, events=("start", "end")):
        if event == "start" and node.tag == qn("w:commentRangeStart"):
            identifier = node.get(qn("w:id"), "")
            active.append(identifier)
            fragments.setdefault(identifier, [])
        elif event == "start" and node.tag == qn("w:t") and active and node.text:
            for identifier in active:
                fragments[identifier].append(node.text)
        elif event == "end" and node.tag == qn("w:commentRangeEnd"):
            identifier = node.get(qn("w:id"), "")
            if identifier in active:
                active.remove(identifier)
    for identifier, values in fragments.items():
        anchors[identifier] = normalized("".join(values))

    checks = []
    evidence = []
    for phrase, required_terms in COMMENT_CHECKS.items():
        target = normalized(phrase)
        identifiers = [identifier for identifier, text in anchors.items() if text == target]
        ok = len(identifiers) == 1 and all(term in comment_text.get(identifiers[0], "") for term in required_terms)
        checks.append(ok)
        evidence.append(f"{target[:30]}={ok}")
    exact_count = len(comment_text) == 3
    return (sum(checks) + int(exact_count)) / 4, f"批注={len(comment_text)}，" + "；".join(evidence)


def navigation_score(output) -> tuple[float, str]:
    root = output.element
    names = root.xpath("//w:bookmarkStart/@w:name")
    anchors = root.xpath("//w:hyperlink/@w:anchor")
    checks = [name in names and name in anchors for name in BOOKMARKS]
    return sum(checks) / len(checks), f"书签={sorted(set(names) & set(BOOKMARKS))}，内链={sorted(set(anchors) & set(BOOKMARKS))}"


def native_structure(output) -> tuple[float, str]:
    table = find_matrix(output)
    matrix_native = table is not None
    borders = False
    repeated_header = False
    nonsplit = False
    if table is not None:
        borders = bool(table._tbl.xpath("./w:tblPr/w:tblBorders | .//w:tcPr/w:tcBorders"))
        repeated_header = bool(table.rows[0]._tr.xpath("./w:trPr/w:tblHeader"))
        nonsplit = all(bool(row._tr.xpath("./w:trPr/w:cantSplit")) for row in table.rows[1:])
    paragraphs = direct_body_paragraphs(output)
    heading_map = {xml_text(p): style_name(output, p) for p in paragraphs if xml_text(p) in {"Policy Integration Review", "Scope Decision", "Research Activity Control Matrix", "Editorial Status"}}
    heading_ok = heading_map == {
        "Policy Integration Review": "Heading 1",
        "Scope Decision": "Heading 2",
        "Research Activity Control Matrix": "Heading 2",
        "Editorial Status": "Heading 2",
    }
    sections_ok = len(output.sections) == 2 and str(output.sections[1].start_type).startswith("NEW_PAGE")
    header_ok = len(output.sections) == 2 and "College of Education — Policy Integration Review" in " ".join(p.text for p in output.sections[1].header.paragraphs)
    checks = [matrix_native, borders, repeated_header, nonsplit, heading_ok, sections_ok, header_ok]
    return sum(checks) / len(checks), f"原生表={matrix_native}，边框={borders}，表头={repeated_header}，不断行={nonsplit}，标题={heading_ok}，分节={sections_ok}，页眉={header_ok}"


def render_check(output_path: Path, source_path: Path) -> tuple[float, str]:
    dangerous: list[str] = []
    with zipfile.ZipFile(output_path) as package:
        names = package.namelist()
        dangerous = [name for name in names if "vbaProject" in name or "/embeddings/" in name or name.endswith(".bin")]
    no_new_external = external_targets(output_path) == external_targets(source_path)
    media_ok = media_hashes(output_path) == media_hashes(source_path)
    office = shutil.which("libreoffice") or shutil.which("soffice")
    if office is None:
        return 0.0, "未找到 LibreOffice"
    try:
        with tempfile.TemporaryDirectory(prefix="syllabus-policy-render-") as temporary:
            subprocess.run(
                [office, "--headless", "--convert-to", "pdf", "--outdir", temporary, str(output_path)],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
            )
            pdf = Path(temporary) / f"{output_path.stem}.pdf"
            text = subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True, errors="replace", timeout=60)
            pages = text.split("\f")
            if pages and not pages[-1].strip():
                pages.pop()
            lengths = [len(normalized(page)) for page in pages]
            render_ok = 8 <= len(pages) <= 14 and all(length >= 80 for length in lengths)
            headings_ok = all(term in text for term in ("ACADEMIC INTEGRITY FOR RESEARCH", "Policy Integration Review", "Research Activity Control Matrix"))
    except Exception as error:
        return 0.0, f"LibreOffice 渲染失败: {error}"
    checks = [not dangerous, no_new_external, media_ok, render_ok, headings_ok]
    return sum(checks) / len(checks), f"危险部件={dangerous}，新外链={not no_new_external}，媒体={media_ok}，页数={len(pages)}，逐页字符={lengths}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()

    inputs = source_directory()
    master_path = inputs / "sample_syllabus.docx"
    policy_path = inputs / "academic_integrity.docx"
    output_path = args.artifact_dir / OUTPUT_NAME
    files = sorted(path.name for path in args.artifact_dir.iterdir() if path.is_file()) if args.artifact_dir.is_dir() else []
    results: list[dict] = []
    artifact_hashes: list[str] = []

    try:
        expected_inputs = sha256(master_path) == MASTER_SHA256 and sha256(policy_path) == POLICY_SHA256
        output_valid = valid_docx(output_path)
        exact_delivery = files == [OUTPUT_NAME]
        r1 = float(expected_inputs and output_valid and exact_delivery)
        results.append(criterion("R001", r1, f"输入快照={expected_inputs}，DOCX={output_valid}，交付列表={files}"))
        if not output_valid:
            raise ValueError("预期 DOCX 缺失或损坏")
        artifact_hashes.append(sha256(output_path))
        master = Document(master_path)
        policy = Document(policy_path)
        output = Document(output_path)

        r2, evidence = protect_master(master_path, output_path, master, output)
        results.append(criterion("R002", r2, evidence))
        r3, r4, evidence = tracked_policy(policy, output)
        results.append(criterion("R003", r3, evidence))
        results.append(criterion("R004", r4, evidence))
        r5, evidence = scope_decision(output)
        results.append(criterion("R005", r5, evidence))
        r6, rows = matrix_score(output)
        results.append(criterion("R006", r6, f"表行={rows}"))
        r7, evidence = comments_score(output_path)
        results.append(criterion("R007", r7, evidence))
        r8, evidence = navigation_score(output)
        results.append(criterion("R008", r8, evidence))
        r9, evidence = native_structure(output)
        results.append(criterion("R009", r9, evidence))
        r10, evidence = render_check(output_path, master_path)
        results.append(criterion("R010", r10, evidence))
    except Exception as error:
        if not results:
            results.append(criterion("R001", 0.0, f"读取失败: {error}"))
        elif len(results) == 1:
            results[0]["evidence"] += f"；后续读取失败={error}"

    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(SPECS[key][0] * values.get(key, 0.0) for key in SPECS) / 100
    hurdle_failures = [key for key, (_, kind, _) in SPECS.items() if kind == "hurdle" and values.get(key, 0.0) < 1.0]
    caps = [SPECS[key][2] for key in hurdle_failures if SPECS[key][2] is not None]
    if caps:
        reward = min(reward, min(caps))
    reward = round(max(0.0, min(1.0, reward)), 6)
    passed = reward >= 0.8 and not hurdle_failures
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": artifact_hashes,
        "criterion_results": results,
        "reward": reward,
        "passed": passed,
        "hurdle_failures": hurdle_failures,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
