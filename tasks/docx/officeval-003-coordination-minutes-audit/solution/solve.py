#!/usr/bin/env python3
"""本题专用：整理部门协调会记录并制作证据回链核查附录。"""

from __future__ import annotations

import argparse
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.shared import Cm, Pt
from lxml import etree


OUTPUT_NAME = "coordination_meeting_minutes_audit.docx"
SOURCE_NAME = "coordination_meeting_minutes_source.docx"
CHANGE_DATE = datetime(2026, 9, 8, 2, 0, tzinfo=timezone.utc).isoformat()
COMMENT_TEXT = "根据本节列明的5条偏差记录核对。"
COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def compact(text: str) -> str:
    """源记录里只有排版空格；正文整理时去掉它们，避免把断行空格当事实。"""
    return re.sub(r"\s+", "", text or "")


def source_table_lines(path: Path) -> list[str]:
    """按真实 XML 顺序递归读取顶层表及其内嵌表，不依赖 python-docx 的浅层 cell.text。"""
    with zipfile.ZipFile(path) as package:
        root = etree.fromstring(package.read("word/document.xml"))
    tables = root.xpath("//w:body/w:tbl", namespaces=W_NS)
    if len(tables) != 1:
        raise ValueError(f"源记录应有 1 个正文顶层表，实际 {len(tables)}")

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
        raise ValueError("无法从源记录拆出完整的 5 条偏差与 3 条整改措施")
    cuts = [0] + [position for position, _ in positions] + [len(text)]
    return [text[cuts[index]:cuts[index + 1]] for index in range(len(cuts) - 1) if text[cuts[index]:cuts[index + 1]]]


def logical_entries(path: Path) -> list[str]:
    """合并排版断行，并把源文件挤在同一单元格的偏差记录拆回语义条目。"""
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
    if len(entries) < 110 or entries[-1] != "许承柏：今天会议议程已全部结束，散会！":
        raise ValueError(f"源记录语义条目异常：{len(entries)}")
    return entries


def extract_deviations(entries: list[str]) -> list[dict[str, str]]:
    pattern = re.compile(
        r"^(?P<no>[1-5])、(?P<date>2026年\d+月\d+日)(?P<time>\d{1,2}:\d{2}至\d{1,2}:\d{2})?，?"
        r"(?P<context>.+?)。责任人(?P<owner>[^，]+)，按照作业质量管理办法定为(?P<class>一般偏差1件)。$"
    )
    rows = []
    for entry in entries:
        match = pattern.match(entry)
        if not match:
            continue
        row = match.groupdict()
        row["time"] = row["time"] or "未记录"
        rows.append(row)
    if [row["no"] for row in rows] != list("12345"):
        raise ValueError(f"偏差记录不连续：{[row['no'] for row in rows]}")
    return rows


def extract_decisions(entries: list[str]) -> list[dict[str, str]]:
    start = entries.index("四、研究后续事项") + 1
    tail = entries[start:]
    title_indices = [index for index, value in enumerate(tail) if re.match(r"^[1-8]、", value)]
    if len(title_indices) != 8:
        raise ValueError(f"后续事项应为 8 项，实际 {len(title_indices)}")
    rows = []
    for position, begin in enumerate(title_indices):
        end = title_indices[position + 1] if position + 1 < len(title_indices) else len(tail) - 1
        block = tail[begin:end]
        number = str(position + 1)
        if not block[0].startswith(number + "、"):
            raise ValueError("后续事项编号与顺序不一致")
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
            raise ValueError(f"议题 {number} 的正式表决不是 4 条：{votes}")
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


def visual_width(char: str) -> float:
    code = ord(char)
    if code < 128:
        return 0.55
    if 0x2000 <= code <= 0x206F:
        return 0.55
    return 1.0


def wrap_for_ruled_rows(text: str, limit: float = 41.5) -> list[str]:
    """把长条目拆成不会在 A4 正文单元格中再次自动换行的可编辑文本行。"""
    lines: list[str] = []
    current = ""
    width = 0.0
    for char in text:
        increment = visual_width(char)
        if current and width + increment > limit:
            if char in "，。；：、！？）》】」』":
                current += char
                lines.append(current)
                current = ""
                width = 0.0
            else:
                lines.append(current)
                current = char
                width = increment
        else:
            current += char
            width += increment
    if current:
        lines.append(current)
    return lines


def clear_paragraph(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def add_text_run(paragraph, text: str, *, bold: bool = False, size: float = 10.5) -> None:
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Noto Sans CJK SC"
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Noto Sans CJK SC")


def plain_run(text: str, *, deleted: bool = False):
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Noto Sans CJK SC")
    fonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "21")
    properties.extend([fonts, size])
    run.append(properties)
    node = OxmlElement("w:delText" if deleted else "w:t")
    node.text = text
    run.append(node)
    return run


def add_bookmarked_text(paragraph, text: str, name: str, identifier: int, *, bold: bool = False) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    paragraph._p.append(start)
    add_text_run(paragraph, text, bold=bold)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.append(end)


def add_count_revision(paragraph, text: str) -> None:
    before, marker, after = text.partition("若干")
    if not marker:
        raise ValueError("未找到需修订的‘若干’")
    paragraph._p.append(plain_run(before))
    deleted = OxmlElement("w:del")
    deleted.set(qn("w:id"), "901")
    deleted.set(qn("w:author"), "会议记录核查员")
    deleted.set(qn("w:date"), CHANGE_DATE)
    deleted.append(plain_run("若干", deleted=True))
    paragraph._p.append(deleted)

    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:id"), "902")
    inserted.set(qn("w:author"), "会议记录核查员")
    inserted.set(qn("w:date"), CHANGE_DATE)
    start = OxmlElement("w:commentRangeStart")
    start.set(qn("w:id"), "1")
    inserted.append(start)
    inserted.append(plain_run("5"))
    end = OxmlElement("w:commentRangeEnd")
    end.set(qn("w:id"), "1")
    inserted.append(end)
    reference_run = OxmlElement("w:r")
    reference_props = OxmlElement("w:rPr")
    reference_style = OxmlElement("w:rStyle")
    reference_style.set(qn("w:val"), "CommentReference")
    reference_props.append(reference_style)
    reference_run.append(reference_props)
    reference = OxmlElement("w:commentReference")
    reference.set(qn("w:id"), "1")
    reference_run.append(reference)
    inserted.append(reference_run)
    paragraph._p.append(inserted)
    paragraph._p.append(plain_run(after))


def add_comments_part(document: Document) -> None:
    root = OxmlElement("w:comments")
    comment = OxmlElement("w:comment")
    comment.set(qn("w:id"), "1")
    comment.set(qn("w:author"), "会议记录核查员")
    comment.set(qn("w:initials"), "核")
    comment.set(qn("w:date"), CHANGE_DATE)
    paragraph = OxmlElement("w:p")
    paragraph.append(plain_run(COMMENT_TEXT))
    comment.append(paragraph)
    root.append(comment)
    part = Part(
        PackURI("/word/comments.xml"),
        COMMENTS_CONTENT_TYPE,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes"),
        document.part.package,
    )
    document.part.relate_to(part, COMMENTS_REL_TYPE)


def set_repeat_header(row) -> None:
    properties = row._tr.get_or_add_trPr()
    if properties.find(qn("w:tblHeader")) is None:
        properties.append(OxmlElement("w:tblHeader"))


def set_no_split(row) -> None:
    properties = row._tr.get_or_add_trPr()
    if properties.find(qn("w:cantSplit")) is None:
        properties.append(OxmlElement("w:cantSplit"))


def set_exact_height(row, value: int = 500) -> None:
    properties = row._tr.get_or_add_trPr()
    height = properties.find(qn("w:trHeight"))
    if height is None:
        height = OxmlElement("w:trHeight")
        properties.append(height)
    height.set(qn("w:val"), str(value))
    height.set(qn("w:hRule"), "exact")


def set_cell_borders(cell, *, ruled: bool) -> None:
    properties = cell._tc.get_or_add_tcPr()
    borders = properties.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single" if ruled and edge == "bottom" else "nil")
        if ruled and edge == "bottom":
            node.set(qn("w:sz"), "4")
            node.set(qn("w:color"), "000000")
            node.set(qn("w:space"), "0")


def set_grid_borders(cell) -> None:
    properties = cell._tc.get_or_add_tcPr()
    borders = properties.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:color"), "7F8C8D")
        node.set(qn("w:space"), "0")


def set_table_layout(table, width_dxa: int) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    properties = table._tbl.tblPr
    layout = properties.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        properties.append(layout)
    layout.set(qn("w:type"), "fixed")
    width = properties.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        properties.append(width)
    width.set(qn("w:w"), str(width_dxa))
    width.set(qn("w:type"), "dxa")


def shade(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    node = properties.find(qn("w:shd"))
    if node is None:
        node = OxmlElement("w:shd")
        properties.append(node)
    node.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, *, bold: bool = False, size: float = 8.0) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    add_text_run(paragraph, text, bold=bold, size=size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_internal_link(cell, target: str, label: str, *, size: float = 8.0) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), target)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    font_size = OxmlElement("w:sz")
    font_size.set(qn("w:val"), str(int(size * 2)))
    properties.extend([color, underline, font_size])
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_field(paragraph, instruction: str) -> None:
    def field_run(node):
        run = OxmlElement("w:r")
        properties = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:ascii"), "Noto Sans CJK SC")
        fonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
        size = OxmlElement("w:sz")
        size.set(qn("w:val"), "16")
        properties.extend([fonts, size])
        run.append(properties)
        run.append(node)
        paragraph._p.append(run)

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin.set(qn("w:dirty"), "true")
    instruction_node = OxmlElement("w:instrText")
    instruction_node.set(qn("xml:space"), "preserve")
    instruction_node.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    result = OxmlElement("w:t")
    result.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for node in (begin, instruction_node, separate, result, end):
        field_run(node)


def configure_header_footer(section, header_text: str) -> None:
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    header = section.header.paragraphs[0]
    clear_paragraph(header)
    for child in list(header._p.get_or_add_pPr()):
        header._p.get_or_add_pPr().remove(child)
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_text_run(header, header_text, size=8.0)
    footer = section.footer.paragraphs[0]
    clear_paragraph(footer)
    for child in list(footer._p.get_or_add_pPr()):
        footer._p.get_or_add_pPr().remove(child)
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.line_spacing = 1.0
    add_text_run(footer, "第 ", size=8.0)
    add_field(footer, "PAGE")
    add_text_run(footer, " 页 / 共 ", size=8.0)
    add_field(footer, "NUMPAGES")
    add_text_run(footer, " 页", size=8.0)


def build_transcript(document: Document, entries: list[str], deviations: list[dict[str, str]], decisions: list[dict[str, str]]):
    source_table = document.tables[0]
    source_table._element.getparent().remove(source_table._element)
    for paragraph in list(document.paragraphs)[::-1]:
        if paragraph.text.strip():
            break
        if paragraph._element.getparent() is document._element.body:
            paragraph._element.getparent().remove(paragraph._element)

    table = document.add_table(rows=0, cols=1)
    set_table_layout(table, 9637)
    deviation_bookmarks = {next(entry for entry in entries if entry.startswith(row["no"] + "、2026年")): f"deviation_{row['no']}" for row in deviations}
    decision_bookmarks = {row["title"]: f"decision_{row['no']}" for row in decisions}
    bookmarks = {**deviation_bookmarks, **decision_bookmarks}
    bookmark_id = 1000

    for entry in entries:
        segments = wrap_for_ruled_rows(entry)
        for segment_index, segment in enumerate(segments):
            row = table.add_row()
            set_no_split(row)
            set_exact_height(row)
            cell = row.cells[0]
            set_cell_borders(cell, ruled=True)
            cell.width = Cm(17.0)
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.left_indent = Pt(0)
            paragraph.paragraph_format.first_line_indent = Pt(0)
            is_heading = bool(re.match(r"^[一二三四]、|^（[一二]）", entry))
            if segment_index == 0 and entry.startswith("二月份作业偏差共发生若干件"):
                add_count_revision(paragraph, segment)
            elif segment_index == 0 and entry in bookmarks:
                add_bookmarked_text(paragraph, segment, bookmarks[entry], bookmark_id, bold=is_heading or entry in decision_bookmarks)
                bookmark_id += 1
            else:
                add_text_run(paragraph, segment, bold=is_heading or (entry in decision_bookmarks and segment_index == 0))
            if is_heading:
                shade(cell, "D9EAF7")
    return table


def format_audit_table(table, headers: list[str], widths_cm: list[float]) -> None:
    set_table_layout(table, int(sum(widths_cm) / 2.54 * 1440))
    for index, width in enumerate(widths_cm):
        for cell in table.columns[index].cells:
            cell.width = Cm(width)
    for row_index, row in enumerate(table.rows):
        set_no_split(row)
        if row_index == 0:
            set_repeat_header(row)
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_grid_borders(cell)
    for cell, header in zip(table.rows[0].cells, headers):
        set_cell_text(cell, header, bold=True, size=8.0)
        shade(cell, "1F4E78")
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = None
            color = OxmlElement("w:color")
            color.set(qn("w:val"), "FFFFFF")
            run._r.get_or_add_rPr().append(color)


def add_heading(document: Document, text: str, level: int = 1) -> None:
    style_name = f"Audit Heading {level}"
    try:
        style = document.styles[style_name]
    except KeyError:
        style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Noto Sans CJK SC"
        style.font.bold = True
        style.font.size = Pt(15 if level == 1 else 11)
    paragraph = document.add_paragraph(text, style=style)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level - 1))
    paragraph._p.get_or_add_pPr().append(outline)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    for run in paragraph.runs:
        run.font.name = "Noto Sans CJK SC"
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), "Noto Sans CJK SC")


def build_appendix(document: Document, deviations: list[dict[str, str]], decisions: list[dict[str, str]]) -> None:
    section = document.add_section(WD_SECTION.NEW_PAGE)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)
    section.top_margin = Cm(1.4)
    section.bottom_margin = Cm(1.4)
    configure_header_footer(section, "会议记录核查附录｜2026-03-18")

    add_heading(document, "会议记录核查附录", 1)
    add_heading(document, "核查摘要", 2)
    summary = document.add_table(rows=5, cols=2)
    summary_values = [
        ("会议日期", "2026年3月18日"),
        ("作业偏差", "5条"),
        ("后续事项", "8项"),
        ("正式表决", "8项，均为4票同意"),
        ("责任人与完成时限", "8项均未明确"),
    ]
    for row, values in zip(summary.rows, summary_values):
        set_no_split(row)
        for cell, value in zip(row.cells, values):
            set_cell_text(cell, value, bold=cell is row.cells[0], size=9.0)
            set_grid_borders(cell)
        shade(row.cells[0], "D9EAF7")
    set_table_layout(summary, 15000)

    add_heading(document, "作业偏差核查表", 2)
    headers = ["回链", "日期", "时段", "检查场景与发现", "责任人", "定性"]
    deviation_table = document.add_table(rows=1, cols=len(headers))
    format_audit_table(deviation_table, headers, [1.5, 2.7, 2.8, 15.0, 2.2, 2.5])
    for source in deviations:
        row = deviation_table.add_row()
        set_no_split(row)
        add_internal_link(row.cells[0], f"deviation_{source['no']}", f"偏差{source['no']}")
        set_grid_borders(row.cells[0])
        values = [source["date"], source["time"], source["context"], source["owner"], source["class"]]
        for cell, value in zip(row.cells[1:], values):
            set_cell_text(cell, value, size=7.5)
            set_grid_borders(cell)

    add_heading(document, "后续事项决议台账", 2)
    headers = ["回链", "事项", "提报人", "事实依据", "正式表决", "最终意见", "责任人", "完成时限"]
    decision_table = document.add_table(rows=1, cols=len(headers))
    format_audit_table(decision_table, headers, [1.5, 5.0, 1.7, 7.2, 4.0, 3.8, 1.6, 1.8])
    for source in decisions:
        row = decision_table.add_row()
        set_no_split(row)
        add_internal_link(row.cells[0], f"decision_{source['no']}", f"议题{source['no']}")
        set_grid_borders(row.cells[0])
        values = [source["subject"], source["proposer"], source["basis"], source["vote"], source["final"], "未明确", "未明确"]
        for cell, value in zip(row.cells[1:], values):
            set_cell_text(cell, value, size=7.0)
            set_grid_borders(cell)


def build(input_dir: Path, output_dir: Path) -> Path:
    source_path = input_dir / SOURCE_NAME
    entries = logical_entries(source_path)
    deviations = extract_deviations(entries)
    decisions = extract_decisions(entries)

    document = Document(source_path)
    settings = document.settings.element
    if settings.find(qn("w:trackRevisions")) is None:
        settings.insert(0, OxmlElement("w:trackRevisions"))
    update_fields = settings.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        settings.append(update_fields)
    update_fields.set(qn("w:val"), "true")
    first = document.sections[0]
    first.orientation = WD_ORIENT.PORTRAIT
    first.page_width = Cm(21.0)
    first.page_height = Cm(29.7)
    first.left_margin = Cm(1.5)
    first.right_margin = Cm(1.5)
    first.top_margin = Cm(1.5)
    first.bottom_margin = Cm(1.5)
    configure_header_footer(first, "部门协调会记录｜整理稿")

    build_transcript(document, entries, deviations, decisions)
    add_comments_part(document)
    build_appendix(document, deviations, decisions)

    document.core_properties.title = "部门协调会记录整理与核查"
    document.core_properties.subject = "2026年3月18日部门协调会受控整理与证据回链核查"
    document.core_properties.keywords = "会议记录,作业偏差,后续事项,受控修订"
    document.core_properties.author = "会议记录核查员"
    document.core_properties.last_modified_by = "会议记录核查员"
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / OUTPUT_NAME
    document.save(target)
    return target


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    arguments = parser.parse_args()
    build(arguments.input_dir, arguments.output_dir)


if __name__ == "__main__":
    main()
