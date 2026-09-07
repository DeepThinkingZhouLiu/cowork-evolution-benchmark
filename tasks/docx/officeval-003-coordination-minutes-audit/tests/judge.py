#!/usr/bin/env python3
"""部门协调会记录整理与核查任务的确定性 Judge。"""

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


TASK_ID = "cowork-evo/officeval-003-coordination-minutes-audit"
OUTPUT_NAME = "coordination_meeting_minutes_audit.docx"
SOURCE_NAME = "coordination_meeting_minutes_source.docx"
REFERENCE_NAME = "ruled_table_reference.png"
RUBRIC_VERSION = "officeval-003-coordination-minutes-audit-v0.3"
JUDGE_VERSION = "officeval-003-coordination-minutes-audit-deterministic-v0.3"
SOURCE_SHA256 = "3c0d0060f17a8b4177c5ed18c6a54502c7a03d3ac6cfdf96b8dadffaaa597344"
REFERENCE_SHA256 = "9da1468839b89dba955a1f50917281e8bd6af5f2863eca2b60fd192a25ba4dc7"
COMMENT_TEXT = "根据本节列明的5条偏差记录核对。"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

SPECS = {
    "R001": (8, "hurdle", 0.0),
    "R002": (12, "hurdle", 0.35),
    "R003": (16, "hurdle", 0.55),
    "R004": (14, "hurdle", 0.65),
    "R005": (14, "hurdle", 0.70),
    "R006": (17, "hurdle", 0.75),
    "R007": (7, "positive", None),
    "R008": (6, "positive", None),
    "R009": (4, "hurdle", 0.78),
    "R010": (2, "positive", None),
}


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_directory() -> Path:
    mounted = Path("/app/input_files")
    if (mounted / SOURCE_NAME).is_file():
        return mounted
    return Path(__file__).resolve().parents[1] / "data/input_files"


def source_table_lines(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    tables = root.xpath("//w:body/w:tbl", namespaces=W_NS)
    if len(tables) != 1:
        raise ValueError(f"冻结源文顶层表数量异常：{len(tables)}")

    def walk(node):
        tag = etree.QName(node).localname
        if tag == "p":
            yield "".join(node.xpath(".//w:t/text()", namespaces=W_NS))
        elif tag == "tbl":
            for row in node.xpath("./w:tr", namespaces=W_NS):
                for cell in row.xpath("./w:tc", namespaces=W_NS):
                    for child in cell:
                        if etree.QName(child).localname in {"p", "tbl"}:
                            yield from walk(child)

    return list(walk(tables[0]))


def split_compound_deviation_section(text: str) -> list[str]:
    markers = [
        "1、2026年", "2、2026年", "3、2026年", "4、2026年", "5、2026年",
        "问题分析：", "整改措施：", "1、加强遵章", "2、开展交心", "3、管理人员增加",
    ]
    positions = sorted((text.index(marker), marker) for marker in markers if marker in text)
    if len(positions) != len(markers):
        raise ValueError("冻结源文无法拆出完整偏差区间")
    cuts = [0] + [position for position, _ in positions] + [len(text)]
    return [text[cuts[index]:cuts[index + 1]] for index in range(len(cuts) - 1) if text[cuts[index]:cuts[index + 1]]]


def logical_entries(path: Path) -> list[str]:
    raw = [compact(value) for value in source_table_lines(path) if compact(value)]
    expanded: list[str] = []
    for value in raw:
        if value.startswith("二月份作业偏差共发生若干件"):
            expanded.extend(split_compound_deviation_section(value))
        else:
            expanded.append(value)
    boundary = re.compile(
        r"^(?:[一二三四]、|（[一二]）|\d+[.、]|二月份作业偏差|许承柏：|许承柏领学：|陈予宁：|侯景澜：|林若棠：|魏预：|周言朔：|最终意见：|问题分析：|整改措施：)"
    )
    entries: list[str] = []
    for value in expanded:
        if not entries or boundary.match(value):
            entries.append(value)
        else:
            entries[-1] += value
    if len(entries) < 110:
        raise ValueError(f"冻结源文语义条目不足：{len(entries)}")
    return entries


def extract_deviations(entries: list[str]) -> list[dict[str, str]]:
    pattern = re.compile(
        r"^(?P<no>[1-5])、(?P<date>2026年\d+月\d+日)(?P<time>\d{1,2}:\d{2}至\d{1,2}:\d{2})?，?"
        r"(?P<context>.+?)。责任人(?P<owner>[^，]+)，按照作业质量管理办法定为(?P<class>一般偏差1件)。$"
    )
    rows = []
    for entry in entries:
        match = pattern.match(entry)
        if match:
            row = match.groupdict()
            row["time"] = row["time"] or "未记录"
            rows.append(row)
    if [row["no"] for row in rows] != list("12345"):
        raise ValueError("冻结源文偏差记录不连续")
    return rows


def extract_decisions(entries: list[str]) -> list[dict[str, str]]:
    start = entries.index("四、研究后续事项") + 1
    tail = entries[start:]
    indices = [index for index, value in enumerate(tail) if re.match(r"^[1-8]、", value)]
    if len(indices) != 8:
        raise ValueError("冻结源文后续事项不是 8 项")
    rows = []
    for position, begin in enumerate(indices):
        end = indices[position + 1] if position + 1 < len(indices) else len(tail) - 1
        block = tail[begin:end]
        number = str(position + 1)
        proposal = next(value for value in block[1:] if re.match(r"^(魏预|周言朔)：", value))
        proposer, basis = proposal.split("：", 1)
        vote_marker = next(index for index, value in enumerate(block) if "现在进行表决" in value)
        final_index = next(index for index, value in enumerate(block) if value.startswith("最终意见："))
        votes = []
        for value in block[vote_marker + 1:final_index]:
            match = re.match(r"^([^：]+)：同意。$", value)
            if match:
                votes.append(match.group(1))
        if len(votes) != 4:
            raise ValueError(f"冻结源文议题 {number} 表决条目异常")
        rows.append({
            "no": number,
            "title": block[0],
            "subject": block[0].split("、", 1)[1],
            "proposer": proposer,
            "basis": basis,
            "vote": "4票同意：" + "、".join(votes),
            "final": block[final_index].split("：", 1)[1],
        })
    return rows


def item(identifier: str, fraction: float, evidence: str) -> dict:
    fraction = max(0.0, min(1.0, fraction))
    return {
        "criterion_id": identifier,
        "score": fraction,
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"passed": fraction == 1.0, "weight": SPECS[identifier][0], "type": SPECS[identifier][1]},
    }


def valid_docx(path: Path) -> bool:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as package:
            return "word/document.xml" in package.namelist() and "[Content_Types].xml" in package.namelist()
    except Exception:
        return False


def xml_text(element) -> str:
    try:
        values = element.xpath(".//w:t/text()", namespaces=W_NS)
    except TypeError:
        values = element.xpath(".//w:t/text()")
    return "".join(values)


def cell_text(cell) -> str:
    return xml_text(cell._tc)


def body_prefix(root) -> str:
    pieces = []
    for child in root.xpath("//w:body/*", namespaces=W_NS):
        if etree.QName(child).localname == "tbl":
            break
        if etree.QName(child).localname == "p":
            pieces.append(xml_text(child))
    return compact("".join(pieces))


def protected_score(source_path: Path, output_path: Path, output: Document) -> tuple[float, str]:
    with zipfile.ZipFile(source_path) as package:
        source_root = etree.fromstring(package.read("word/document.xml"))
    with zipfile.ZipFile(output_path) as package:
        output_root = etree.fromstring(package.read("word/document.xml"))
    prefix_exact = body_prefix(source_root) == body_prefix(output_root)
    prefix = body_prefix(output_root)
    fields = [
        "部门协调会", "2026年3月18日", "主持人:许承柏", "记录人:林若棠",
        "地点:综合楼二层会议室", "应出席数:5", "实出席数:5",
        "列席人员:陈予宁、侯景澜、魏预、周言朔", "会议内容:",
    ]
    field_hits = sum(compact(value) in prefix for value in fields)
    portrait = bool(output.sections and int(output.sections[0].orientation) == 0)
    checks = [prefix_exact, field_hits == len(fields), portrait]
    return sum(checks) / len(checks), f"顶部逐字={prefix_exact}，字段={field_hits}/{len(fields)}，正文纵向={portrait}"


def comment_anchor_text(document_root, identifier: str) -> str:
    active = False
    parts = []
    for event, node in etree.iterwalk(document_root, events=("start", "end")):
        if event == "start" and node.tag == qn("w:commentRangeStart") and node.get(qn("w:id"), "") == identifier:
            active = True
        elif event == "start" and node.tag == qn("w:t") and active and node.text:
            parts.append(node.text)
        elif event == "end" and node.tag == qn("w:commentRangeEnd") and node.get(qn("w:id"), "") == identifier:
            active = False
    return "".join(parts)


def fidelity_score(source_path: Path, output_path: Path, entries: list[str]) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
        settings = etree.fromstring(package.read("word/settings.xml"))
        tables = root.xpath("//w:body/w:tbl", namespaces=W_NS)
        if not tables:
            return 0.0, "缺少会议正文表"
        transcript = tables[0]
        actual = compact(xml_text(transcript))
        expected_entries = [value.replace("若干", "5", 1) if value.startswith("二月份作业偏差共发生若干件") else value for value in entries]
        expected = compact("".join(expected_entries))
        coverage = sum(compact(value) in actual for value in expected_entries) / len(expected_entries)
        exact = actual == expected
        deleted = "".join(root.xpath("//w:body/w:tbl[1]//w:del//w:delText/text()", namespaces=W_NS))
        inserted = "".join(root.xpath("//w:body/w:tbl[1]//w:ins//w:t/text()", namespaces=W_NS))
        authors = root.xpath("//w:body/w:tbl[1]//w:del/@w:author | //w:body/w:tbl[1]//w:ins/@w:author", namespaces=W_NS)
        revision_ok = deleted == "若干" and inserted == "5" and authors and all(value == "会议记录核查员" for value in authors)
        track_enabled = bool(settings.xpath("//w:trackRevisions", namespaces=W_NS))
        try:
            comments = etree.fromstring(package.read("word/comments.xml"))
            bodies = {
                node.get(qn("w:id"), ""): compact(xml_text(node))
                for node in comments.xpath("//w:comment", namespaces=W_NS)
            }
            matched = [identifier for identifier, body in bodies.items() if body == compact(COMMENT_TEXT)]
            comment_ok = len(bodies) == 1 and len(matched) == 1 and comment_anchor_text(root, matched[0]) == "5"
        except (KeyError, etree.XMLSyntaxError):
            comment_ok = False
    fraction = (coverage + int(exact) + int(revision_ok) + int(track_enabled) + int(comment_ok)) / 5
    return fraction, f"语义条目={coverage:.3f}，全文精确={exact}，修订={revision_ok}，Track Changes={track_enabled}，批注={comment_ok}"


def visual_width(text: str) -> float:
    width = 0.0
    for char in text:
        code = ord(char)
        width += 0.55 if code < 128 or 0x2000 <= code <= 0x206F else 1.0
    return width


def transcript_structure_score(output_path: Path) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    tables = root.xpath("//w:body/w:tbl", namespaces=W_NS)
    if not tables:
        return 0.0, "缺少正文表"
    table = tables[0]
    rows = table.xpath("./w:tr", namespaces=W_NS)
    one_column = bool(rows) and all(len(row.xpath("./w:tc", namespaces=W_NS)) == 1 for row in rows)
    no_nested = not table.xpath(".//w:tc/w:tbl", namespaces=W_NS)
    nonempty = all(compact(xml_text(row)) for row in rows)
    one_paragraph = all(len(row.xpath("./w:tc/w:p", namespaces=W_NS)) == 1 for row in rows)
    no_breaks = not table.xpath(".//w:br", namespaces=W_NS)
    max_width = max((visual_width(compact(xml_text(row))) for row in rows), default=999.0)
    no_wrap = max_width <= 42.6
    heights = []
    ruled = 0
    no_sides = 0
    nonsplit = 0
    for row in rows:
        height = row.xpath("./w:trPr/w:trHeight", namespaces=W_NS)
        if height and height[0].get(qn("w:hRule")) == "exact":
            try:
                heights.append(int(height[0].get(qn("w:val"), "0")))
            except ValueError:
                heights.append(0)
        nonsplit += int(bool(row.xpath("./w:trPr/w:cantSplit", namespaces=W_NS)))
        cell = row.xpath("./w:tc", namespaces=W_NS)[0] if row.xpath("./w:tc", namespaces=W_NS) else None
        if cell is None:
            continue
        bottoms = cell.xpath("./w:tcPr/w:tcBorders/w:bottom", namespaces=W_NS)
        if bottoms:
            node = bottoms[0]
            ruled += int(node.get(qn("w:val")) == "single" and node.get(qn("w:sz")) == "4" and node.get(qn("w:color"), "").upper() in {"000000", "AUTO"})
        sides = cell.xpath("./w:tcPr/w:tcBorders/w:left | ./w:tcPr/w:tcBorders/w:right", namespaces=W_NS)
        no_sides += int(not sides or all(node.get(qn("w:val")) in {"nil", "none"} for node in sides))
    equal_height = len(heights) == len(rows) and bool(heights) and max(heights) == min(heights) and 480 <= heights[0] <= 520
    fraction = sum([
        int(one_column and no_nested), int(nonempty and one_paragraph), int(no_breaks and no_wrap),
        ruled / len(rows) if rows else 0.0, no_sides / len(rows) if rows else 0.0,
        nonsplit / len(rows) if rows else 0.0, int(equal_height), int(145 <= len(rows) <= 180),
    ]) / 8
    return fraction, (
        f"行={len(rows)}，单列无嵌套={one_column and no_nested}，无空行/单段={nonempty and one_paragraph}，"
        f"不换行={no_breaks and no_wrap}(最大宽={max_width:.1f})，底线={ruled}/{len(rows)}，"
        f"无竖线={no_sides}/{len(rows)}，不拆行={nonsplit}/{len(rows)}，等高={equal_height}"
    )


def exact_table_score(table, headers: list[str], rows: list[list[str]]) -> tuple[float, str]:
    shape = len(table.rows) == len(rows) + 1 and len(table.columns) == len(headers)
    actual_headers = [cell_text(cell) for cell in table.rows[0].cells] if table.rows else []
    header_score = sum(compact(a) == compact(b) for a, b in zip(actual_headers, headers)) / len(headers)
    passed = 0
    total = len(rows) * len(headers)
    for row_index, expected in enumerate(rows, start=1):
        if row_index >= len(table.rows):
            continue
        actual = [cell_text(cell) for cell in table.rows[row_index].cells]
        passed += sum(compact(a) == compact(b) for a, b in zip(actual, expected))
    fraction = (int(shape) + header_score + (passed / total if total else 0.0)) / 3
    return fraction, f"形状={len(table.rows)}x{len(table.columns)}({shape})，表头={header_score:.3f}，字段={passed}/{total}"


def deviation_score(output: Document, deviations: list[dict[str, str]]) -> tuple[float, str]:
    if len(output.tables) < 3:
        return 0.0, "缺少作业偏差核查表"
    headers = ["回链", "日期", "时段", "检查场景与发现", "责任人", "定性"]
    rows = [
        [f"偏差{row['no']}", row["date"], row["time"], row["context"], row["owner"], row["class"]]
        for row in deviations
    ]
    return exact_table_score(output.tables[2], headers, rows)


def decision_score(output: Document, decisions: list[dict[str, str]]) -> tuple[float, str]:
    if len(output.tables) < 4:
        return 0.0, "缺少后续事项决议台账"
    headers = ["回链", "事项", "提报人", "事实依据", "正式表决", "最终意见", "责任人", "完成时限"]
    rows = [
        [f"议题{row['no']}", row["subject"], row["proposer"], row["basis"], row["vote"], row["final"], "未明确", "未明确"]
        for row in decisions
    ]
    return exact_table_score(output.tables[3], headers, rows)


def navigation_score(output_path: Path, deviations: list[dict[str, str]], decisions: list[dict[str, str]]) -> tuple[float, str]:
    with zipfile.ZipFile(output_path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    tables = root.xpath("//w:body/w:tbl", namespaces=W_NS)
    if not tables:
        return 0.0, "缺少正文表"
    expected = [(f"deviation_{row['no']}", row["no"] + "、2026年") for row in deviations]
    expected += [(f"decision_{row['no']}", row["title"]) for row in decisions]
    passed = 0
    details = []
    for name, anchor in expected:
        bookmarks = tables[0].xpath(f".//w:bookmarkStart[@w:name='{name}']", namespaces=W_NS)
        links = root.xpath(f"//w:hyperlink[@w:anchor='{name}']", namespaces=W_NS)
        rows = bookmarks[0].xpath("ancestor::w:tr[1]", namespaces=W_NS) if len(bookmarks) == 1 else []
        anchor_ok = len(rows) == 1 and compact(anchor) in compact(xml_text(rows[0]))
        ok = anchor_ok and len(links) == 1
        passed += int(ok)
        details.append(f"{name}={ok}")
    return passed / len(expected), f"回链={passed}/{len(expected)}；" + ",".join(details)


def footer_fields(section) -> set[str]:
    root = section.footer._element
    fields = {compact(value).upper() for value in root.xpath(".//w:fldSimple/@w:instr")}
    for value in root.xpath(".//w:instrText/text()"):
        instruction = compact(value).upper()
        if instruction.startswith("NUMPAGES"):
            fields.add("NUMPAGES")
        elif instruction.startswith("PAGE"):
            fields.add("PAGE")
    return fields


def native_score(output: Document) -> tuple[float, str]:
    sections = len(output.sections) == 2
    orientations = sections and int(output.sections[0].orientation) == 0 and int(output.sections[1].orientation) == 1
    headers = sections and [compact("".join(p.text for p in section.header.paragraphs)) for section in output.sections] == [
        compact("部门协调会记录｜整理稿"), compact("会议记录核查附录｜2026-03-18")
    ]
    footers = sections and all({"PAGE", "NUMPAGES"}.issubset(footer_fields(section)) for section in output.sections)
    table_shapes = [(len(table.rows), len(table.columns)) for table in output.tables]
    shapes = table_shapes[:4] == [(table_shapes[0][0], 1), (5, 2), (6, 6), (9, 8)] if len(table_shapes) == 4 else False
    repeat = len(output.tables) == 4 and all(output.tables[index].rows[0]._tr.xpath("./w:trPr/w:tblHeader") for index in (2, 3))
    nonsplit = len(output.tables) == 4 and all(
        row._tr.xpath("./w:trPr/w:cantSplit") for index in (2, 3) for row in output.tables[index].rows[1:]
    )
    properties = output.core_properties.title == "部门协调会记录整理与核查" and output.core_properties.subject == "2026年3月18日部门协调会受控整理与证据回链核查"
    checks = [sections, orientations, headers, footers, shapes, repeat, nonsplit, properties]
    return sum(checks) / len(checks), (
        f"分节={sections}，方向={orientations}，页眉={headers}，页脚字段={footers}，"
        f"表形={table_shapes}，重复表头={repeat}，不拆行={nonsplit}，属性={properties}"
    )


def consistency_score(source_path: Path, output: Document, decisions: list[dict[str, str]]) -> tuple[float, str]:
    if len(output.tables) < 4:
        return 0.0, "附录表不完整"
    expected_summary = [
        ["会议日期", "2026年3月18日"], ["作业偏差", "5条"], ["后续事项", "8项"],
        ["正式表决", "8项，均为4票同意"], ["责任人与完成时限", "8项均未明确"],
    ]
    summary = [[cell_text(cell) for cell in row.cells] for row in output.tables[1].rows]
    summary_ok = [[compact(value) for value in row] for row in summary] == [[compact(value) for value in row] for row in expected_summary]
    unknown_ok = all(cell_text(output.tables[3].cell(index, column)) == "未明确" for index in range(1, 9) for column in (6, 7))
    votes_ok = all(compact(cell_text(output.tables[3].cell(index, 4))) == compact(row["vote"]) for index, row in enumerate(decisions, start=1))
    source = Document(source_path)
    allowed = {compact(p.text) for p in source.paragraphs if compact(p.text)} | {
        compact("会议记录核查附录"), compact("核查摘要"), compact("作业偏差核查表"), compact("后续事项决议台账")
    }
    body = [compact(p.text) for p in output.paragraphs if compact(p.text)]
    no_extra_body = all(value in allowed for value in body)
    checks = [summary_ok, unknown_ok, votes_ok, no_extra_body]
    return sum(checks) / len(checks), f"摘要={summary_ok}，未决字段={unknown_ok}，表决一致={votes_ok}，无额外正文={no_extra_body}"


def safety_score(path: Path) -> tuple[float, str]:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        forbidden = [
            name for name in names
            if name.endswith(("vbaProject.bin", ".bin")) or "embeddings/" in name or "activeX/" in name or "altChunk" in name
        ]
        media = [name for name in names if name.startswith("word/media/")]
        external = []
        for name in names:
            if not name.endswith(".rels"):
                continue
            root = etree.fromstring(package.read(name))
            external.extend(node.get("Target", "") for node in root if node.get("TargetMode") == "External")
    safe = not forbidden and not media and not external
    return float(safe), f"禁止部件={forbidden}，媒体={media}，外链={external}"


def render_score(path: Path) -> tuple[float, str]:
    if not shutil.which("libreoffice") or not shutil.which("pdftotext") or not shutil.which("pdfinfo"):
        return 0.0, "缺少 LibreOffice/Poppler"
    with tempfile.TemporaryDirectory(prefix="coordination-minutes-judge-") as temporary:
        root = Path(temporary)
        profile = root / "profile"
        process = subprocess.run(
            ["libreoffice", "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(root), str(path)],
            capture_output=True, text=True, timeout=180, check=False,
        )
        pdf = root / f"{path.stem}.pdf"
        if process.returncode != 0 or not pdf.is_file():
            return 0.0, "LibreOffice PDF 转换失败"
        text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, timeout=60, check=False).stdout
        info = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True, timeout=60, check=False).stdout
        match = re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        rendered = compact(text)
        required = ["部门协调会", "三、作业偏差情况分析", "四、研究后续事项", "会议记录核查附录", "作业偏差核查表", "后续事项决议台账"]
        checks = [7 <= pages <= 10, len(rendered) >= 2500, all(compact(value) in rendered for value in required)]
        return sum(checks) / len(checks), f"页数={pages}，渲染字符={len(rendered)}，关键内容={all(compact(value) in rendered for value in required)}"


def finalize(results: list[dict], output_path: Path | None) -> tuple[dict, float, bool]:
    scores = {result["criterion_id"]: result["score"] for result in results}
    reward = sum(SPECS[identifier][0] * scores.get(identifier, 0.0) for identifier in SPECS) / 100
    caps = [spec[2] for identifier, spec in SPECS.items() if spec[1] == "hurdle" and scores.get(identifier, 0.0) < 1.0 and spec[2] is not None]
    if caps:
        reward = min(reward, min(caps))
    reward = round(reward, 6)
    passed = reward >= 0.8 and all(scores.get(identifier, 0.0) == 1.0 for identifier in ("R001", "R002", "R003", "R004", "R005", "R006", "R009"))
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
    output_path = artifact_dir / OUTPUT_NAME
    delivered = sorted(path.name for path in artifact_dir.iterdir() if path.is_file() and not path.name.startswith("~$")) if artifact_dir.is_dir() else []
    valid = valid_docx(output_path)
    results = [item("R001", float(valid and delivered == [OUTPUT_NAME]), f"文件={delivered}，原生DOCX={valid}")]
    if not valid:
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, "核心 DOCX 缺失或无法解析"))
        return finalize(results, output_path if output_path.is_file() else None)

    source_dir = source_directory()
    source_path = source_dir / SOURCE_NAME
    reference_path = source_dir / REFERENCE_NAME
    hashes = (sha256(source_path), sha256(reference_path))
    if hashes != (SOURCE_SHA256, REFERENCE_SHA256):
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, f"冻结输入校验失败：{hashes}"))
        return finalize(results, output_path)

    entries = logical_entries(source_path)
    deviations = extract_deviations(entries)
    decisions = extract_decisions(entries)
    output = Document(output_path)

    fraction, evidence = protected_score(source_path, output_path, output)
    results.append(item("R002", fraction, evidence))
    fraction, evidence = fidelity_score(source_path, output_path, entries)
    results.append(item("R003", fraction, evidence))
    fraction, evidence = transcript_structure_score(output_path)
    results.append(item("R004", fraction, evidence))
    fraction, evidence = deviation_score(output, deviations)
    results.append(item("R005", fraction, evidence))
    fraction, evidence = decision_score(output, decisions)
    results.append(item("R006", fraction, evidence))
    fraction, evidence = navigation_score(output_path, deviations, decisions)
    results.append(item("R007", fraction, evidence))
    fraction, evidence = native_score(output)
    results.append(item("R008", fraction, evidence))
    fraction, evidence = consistency_score(source_path, output, decisions)
    results.append(item("R009", fraction, evidence))
    render_fraction, render_evidence = render_score(output_path)
    safety_fraction, safety_evidence = safety_score(output_path)
    results.append(item("R010", (render_fraction + safety_fraction) / 2, f"{render_evidence}；{safety_evidence}"))
    return finalize(results, output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    arguments = parser.parse_args()
    try:
        payload, reward, passed = evaluate(arguments.artifact_dir)
    except Exception as exc:
        results = [item("R001", 0.0, f"Judge 读取失败：{type(exc).__name__}: {exc}")]
        for identifier in list(SPECS)[1:]:
            results.append(item(identifier, 0.0, "未评分"))
        payload, reward, passed = finalize(results, None)
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.reward_file.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    arguments.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
