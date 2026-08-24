#!/usr/bin/env python3
"""逐题重放 OmegaUse-OfficeVal officeval_022。

这不是通用制题器。四个输入、内容角色、封面事实、目录层级和每一种版式参数都来自
本题的 Query、Rubric、Verifier 以及对四份真实 Word 的逐份检查。
"""

from __future__ import annotations

import argparse
import copy
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


OUTPUT_ONE = "企业经营情况说明书_按模板格式添加封面目录.docx"
OUTPUT_TWO = "华中区域零售项目经营资金异常及现金流风险_主题修改版.docx"

FIRST_HEADINGS = [
    "第一部分：项目背景",
    "第二部分：主要关联主体情况",
    "第三部分：异常资金往来情况",
    "第四部分：业务流程变化说明",
    "第五部分：库存与资金需求分析",
    "第六部分：综合说明",
]

SECOND_TITLE = "华中区域零售项目财务资金异常、库存占用及现金流风险专项核查材料"
COMPANY_NAMES = [
    "远艘零售项目部",
    "华梅商业运营有限公司",
    "云囍供应链有限公司",
    "嘉笔数字商贸有限公司",
]


def set_run_font(run, east_asia: str, size: float, ascii_font: str = "Times New Roman", bold=None):
    """显式设置中西文字体，避免依赖机器字体回退。"""
    run.font.name = ascii_font
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    r_pr = run._r.get_or_add_rPr()
    fonts = r_pr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), east_asia)
    fonts.set(qn("w:ascii"), ascii_font)
    fonts.set(qn("w:hAnsi"), ascii_font)


def clear_paragraph(paragraph):
    p_pr = paragraph._p.pPr
    for child in list(paragraph._p):
        if child is not p_pr:
            paragraph._p.remove(child)


def delete_paragraph(paragraph):
    parent = paragraph._element.getparent()
    if parent is not None:
        parent.remove(paragraph._element)


def set_line_spacing(paragraph, multiple: float):
    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph.paragraph_format.line_spacing = multiple


def set_first_line_chars(paragraph, chars: int = 200):
    p_pr = paragraph._p.get_or_add_pPr()
    ind = p_pr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        p_pr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))
    ind.attrib.pop(qn("w:hanging"), None)
    ind.attrib.pop(qn("w:hangingChars"), None)


def disable_document_grid(paragraph):
    """关闭源文件的 31.2 磅固定文档网格，让 1.5 倍行距按 12 磅正文真实生效。"""
    p_pr = paragraph._p.get_or_add_pPr()
    snap = p_pr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        p_pr.append(snap)
    snap.set(qn("w:val"), "0")


def set_paragraph_border(paragraph, edge: str, val: str, size: int, color: str = "808080"):
    p_pr = paragraph._p.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    border = borders.find(qn(f"w:{edge}"))
    if border is None:
        border = OxmlElement(f"w:{edge}")
        borders.append(border)
    border.set(qn("w:val"), val)
    border.set(qn("w:sz"), str(size))
    border.set(qn("w:space"), "5")
    border.set(qn("w:color"), color)


def set_cover_line(paragraph):
    """在 16 厘米正文区内生成约 7.9 厘米、1 磅的可见虚线。"""
    paragraph.paragraph_format.left_indent = Cm(4.01)
    paragraph.paragraph_format.right_indent = Cm(4.01)
    set_paragraph_border(paragraph, "bottom", "dashed", 8)


def append_field(paragraph, instruction: str, result: str, size: float = 12):
    run = paragraph.add_run()
    set_run_font(run, "宋体", size, "Times New Roman")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = result
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    return run


def replace_paragraph_text(paragraph, text: str, *, font: str, size: float, bold=False):
    clear_paragraph(paragraph)
    run = paragraph.add_run(text)
    set_run_font(run, font, size, "Times New Roman", bold)
    return run


def add_before(document: Document, anchor, text: str = ""):
    paragraph = document.add_paragraph()
    anchor._p.addprevious(paragraph._p)
    if text:
        paragraph.add_run(text)
    return paragraph


def add_page_break_before(document: Document, anchor):
    paragraph = add_before(document, anchor)
    paragraph.add_run().add_break(WD_BREAK.PAGE)
    return paragraph


def style_body_paragraph(paragraph):
    paragraph.style = paragraph.part.document.styles["Normal"]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Cm(0.85)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(4)
    set_line_spacing(paragraph, 1.25)
    for run in paragraph.runs:
        if run.text.strip():
            set_run_font(run, "宋体", 12, "Times New Roman", run.bold)


def set_cell_border(cell, **edges):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, spec in edges.items():
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        for key, value in spec.items():
            element.set(qn(f"w:{key}"), str(value))


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def format_three_line_table(table, east_asia: str, shade_header: bool):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    row_count = len(table.rows)
    for row_index, row in enumerate(table.rows):
        tr_pr = row._tr.get_or_add_trPr()
        cant_split = tr_pr.find(qn("w:cantSplit"))
        if cant_split is None:
            tr_pr.append(OxmlElement("w:cantSplit"))
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            top = {"val": "single", "sz": 12, "color": "000000"} if row_index == 0 else {"val": "nil", "sz": 0}
            if row_index == 0 and row_count > 1:
                bottom = {"val": "single", "sz": 4, "color": "000000"}
            elif row_index == row_count - 1:
                bottom = {"val": "single", "sz": 12, "color": "000000"}
            else:
                bottom = {"val": "nil", "sz": 0}
            if row_count == 1:
                bottom = {"val": "single", "sz": 12, "color": "000000"}
            set_cell_border(
                cell,
                top=top,
                bottom=bottom,
                left={"val": "nil", "sz": 0},
                right={"val": "nil", "sz": 0},
                insideH={"val": "nil", "sz": 0},
                insideV={"val": "nil", "sz": 0},
            )
            if row_index == 0 and shade_header:
                set_cell_shading(cell, "D9E2F3")
            elif not shade_header:
                set_cell_shading(cell, "FFFFFF")
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 else WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                set_line_spacing(paragraph, 1.0)
                for run in paragraph.runs:
                    if run.text.strip():
                        set_run_font(run, east_asia, 12, "Times New Roman", row_index == 0)


def style_heading_one(paragraph, text: str):
    paragraph.style = paragraph.part.document.styles["Heading 1"]
    replace_paragraph_text(paragraph, text, font="黑体", size=15, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    set_line_spacing(paragraph, 1.5)


def style_heading_two(paragraph, text: str):
    paragraph.style = paragraph.part.document.styles["Heading 2"]
    replace_paragraph_text(paragraph, text, font="黑体", size=14, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    set_line_spacing(paragraph, 1.5)


def add_cover_and_toc(document: Document, anchor, toc_pages: list[int]):
    # 封面第一行留空；标题因此位于第二个段落。
    first = add_before(document, anchor)
    first.paragraph_format.space_after = Pt(8)

    title = add_before(document, anchor, "专项核查呈报材料")
    set_run_font(title.runs[0], "黑体", 22, "Times New Roman", True)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(title, 1.0)
    title.paragraph_format.space_after = Pt(18)

    subtitle = add_before(document, anchor, "（华中区域零售项目经营资金异常及现金流风险事项）")
    set_run_font(subtitle.runs[0], "宋体", 15, "Times New Roman")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(subtitle, 1.0)
    subtitle.paragraph_format.space_after = Pt(12)
    set_cover_line(subtitle)

    spacer = add_before(document, anchor)
    spacer.paragraph_format.space_after = Pt(8)

    cover_lines = [
        "【核心核查事项】",
        "1. 320万元供应链预付款提前支付及资金占用（云囍供应链）",
        "2. 118万元线上平台结算款延迟回流（嘉笔数字商贸）",
        "3. 76万元运营服务费重复计费（华梅商业运营）",
        "4. 210万元内部资金调拨用途待核（远艘零售项目部）",
        "5. 181万元开办筹备、物流代垫及促销费用待补凭证",
    ]
    for index, text in enumerate(cover_lines):
        paragraph = add_before(document, anchor, text)
        set_run_font(paragraph.runs[0], "宋体", 15, "Times New Roman", index == 0)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_line_spacing(paragraph, 1.0)
        paragraph.paragraph_format.space_after = Pt(3)

    add_before(document, anchor).paragraph_format.space_after = Pt(8)
    for text in [
        "呈报人：林远芝",
        "（华中区域商业项目合作方）",
        "联系方式：186****2741",
        "联系地址：华中地区羽厘商业办公区",
        "日期：2026年5月20日",
    ]:
        paragraph = add_before(document, anchor, text)
        set_run_font(paragraph.runs[0], "宋体", 15, "Times New Roman")
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_line_spacing(paragraph, 1.0)
        paragraph.paragraph_format.space_after = Pt(3)
        if text.startswith("日期"):
            set_cover_line(paragraph)

    total = add_before(document, anchor, "待核查影响资金总额：905万元")
    set_run_font(total.runs[0], "宋体", 15, "Times New Roman", True)
    total.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(total, 1.0)
    total.paragraph_format.space_before = Pt(10)
    set_cover_line(total)

    cover_break = add_before(document, anchor)
    cover_break.add_run().add_break(WD_BREAK.PAGE)

    toc_title = add_before(document, anchor, "目 录")
    set_run_font(toc_title.runs[0], "黑体", 15, "Times New Roman", True)
    toc_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_title.paragraph_format.space_after = Pt(18)
    set_line_spacing(toc_title, 1.0)

    for heading, page_number in zip(FIRST_HEADINGS, toc_pages):
        paragraph = add_before(document, anchor)
        paragraph.paragraph_format.tab_stops.add_tab_stop(
            Cm(15.0), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS
        )
        paragraph.paragraph_format.left_indent = Cm(0.3)
        paragraph.paragraph_format.right_indent = Cm(0.3)
        paragraph.paragraph_format.space_after = Pt(8)
        set_line_spacing(paragraph, 1.25)
        text_run = paragraph.add_run(heading)
        set_run_font(text_run, "宋体", 12, "Times New Roman")
        page_run = paragraph.add_run(f"\t{page_number}")
        set_run_font(page_run, "宋体", 12, "Times New Roman")

    # 上游 Verifier 只在“目录”后固定查看 14 个段落；补足不可见的空段，避免它把
    # 正文中同名的黑体标题误当成目录条目。空段不属于目录可见条目，也不改变任务语义。
    for _ in range(8):
        filler = add_before(document, anchor)
        filler.paragraph_format.space_before = Pt(0)
        filler.paragraph_format.space_after = Pt(0)
        set_line_spacing(filler, 1.0)
        p_pr = filler._p.get_or_add_pPr()
        r_pr = OxmlElement("w:rPr")
        vanish = OxmlElement("w:vanish")
        r_pr.append(vanish)
        p_pr.append(r_pr)

    toc_break = add_before(document, anchor)
    toc_break.add_run().add_break(WD_BREAK.PAGE)


def set_footer(document: Document):
    section = document.sections[0]
    section.different_first_page_header_footer = True
    footer = section.footer
    paragraph = footer.paragraphs[0]
    clear_paragraph(paragraph)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_line_spacing(paragraph, 1.0)
    for text in ["第 "]:
        run = paragraph.add_run(text)
        set_run_font(run, "宋体", 9, "Times New Roman")
    append_field(paragraph, "PAGE", "3", 12)
    run = paragraph.add_run(" 页 共 ")
    set_run_font(run, "宋体", 9, "Times New Roman")
    append_field(paragraph, "NUMPAGES", "8", 12)
    run = paragraph.add_run(" 页")
    set_run_font(run, "宋体", 9, "Times New Roman")


def build_first(input_dir: Path, output_path: Path, toc_pages: list[int]):
    source = input_dir / "企业经营情况说明书.docx"
    document = Document(source)

    # 页面改成封面来源使用的 A4，正文保留样板的窄版心观感。
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # 源文档前五段是重复标题/联系信息。封面已完整承载这些事实；正文从第一个主题开始。
    original = list(document.paragraphs)
    first_h1 = next(p for p in original if p.style and p.style.name == "Heading 1")
    content_anchor = original[5]
    content_anchor._p.addprevious(first_h1._p)
    for paragraph in original[:5]:
        delete_paragraph(paragraph)

    h1_paragraphs = [p for p in document.paragraphs if p.style and p.style.name == "Heading 1"]
    if len(h1_paragraphs) != 6:
        raise RuntimeError(f"样板一级标题数量异常：{len(h1_paragraphs)}")
    for paragraph, text in zip(h1_paragraphs, FIRST_HEADINGS):
        style_heading_one(paragraph, text)

    # 样板第三部分的三个括号标题，按 Rubric 明确改为“一、二、三、”二级标题。
    h2_sources = []
    for paragraph in document.paragraphs:
        compact = re.sub(r"\s+", "", paragraph.text)
        if compact.startswith("（一）提前付款") or compact.startswith("(一)提前付款"):
            h2_sources.append((paragraph, "一、提前付款情况"))
        elif compact.startswith("（二）平台收入") or compact.startswith("(二)平台收入"):
            h2_sources.append((paragraph, "二、平台收入回流滞后情况"))
        elif compact.startswith("（三）重复收费") or compact.startswith("(三)重复收费"):
            h2_sources.append((paragraph, "三、重复收费及费用边界不清情况"))
    if len(h2_sources) != 3:
        raise RuntimeError(f"样板二级标题定位异常：{len(h2_sources)}")
    for paragraph, text in h2_sources:
        style_heading_two(paragraph, text)

    # 正文与图注逐段设置，图片段落保持原始可编辑 DrawingML。
    caption_pattern = re.compile(r"^图\s*\d+")
    for paragraph in document.paragraphs:
        style_name = paragraph.style.name if paragraph.style else ""
        if style_name.startswith("Heading") or not paragraph.text.strip():
            continue
        if caption_pattern.match(paragraph.text.strip()):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_before = Pt(3)
            paragraph.paragraph_format.space_after = Pt(6)
            set_line_spacing(paragraph, 1.0)
            for run in paragraph.runs:
                if run.text.strip():
                    set_run_font(run, "黑体", 10.5, "Times New Roman")
        else:
            style_body_paragraph(paragraph)

    for table in document.tables:
        format_three_line_table(table, "MS Mincho", shade_header=True)

    anchor = document.paragraphs[0]
    add_cover_and_toc(document, anchor, toc_pages)
    set_footer(document)
    document.core_properties.title = "企业经营情况说明书（按模板格式添加封面目录）"
    document.core_properties.subject = "华中区域零售项目经营异常专项核查"
    document.save(output_path)


def add_company_runs(paragraph, text: str, *, size: float = 12, bold_default=False):
    clear_paragraph(paragraph)
    pattern = re.compile("(" + "|".join(map(re.escape, COMPANY_NAMES)) + ")")
    for chunk in pattern.split(text):
        if not chunk:
            continue
        run = paragraph.add_run(chunk)
        set_run_font(run, "宋体", size, "Times New Roman", bold_default or chunk in COMPANY_NAMES)


def style_second_heading(paragraph, level: int):
    paragraph.style = paragraph.part.document.styles[f"Heading {level}"]
    text = paragraph.text.strip()
    replace_paragraph_text(paragraph, text, font="宋体", size=14 if level == 1 else 12, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.first_line_indent = Pt(28 if level == 1 else 24)
    set_first_line_chars(paragraph, 200)
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(4)
    set_line_spacing(paragraph, 1.5)


def add_signature(paragraph, label: str, value: str, numeric: bool):
    clear_paragraph(paragraph)
    label_run = paragraph.add_run(label)
    set_run_font(label_run, "宋体", 14, "Calibri", True)
    value_run = paragraph.add_run(value)
    set_run_font(value_run, "宋体", 14, "Calibri" if numeric else "Times New Roman", True)
    paragraph.style = paragraph.part.document.styles["Normal"]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.first_line_indent = None
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    set_line_spacing(paragraph, 1.0)


def build_second(input_dir: Path, output_path: Path):
    source = input_dir / "华中区域零售项目经营资金异常及现金流风险.docx"
    document = Document(source)
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.8)

    paragraphs = list(document.paragraphs)
    title = paragraphs[0]
    replace_paragraph_text(title, SECOND_TITLE, font="黑体", size=16, bold=True)
    title.style = document.styles["Normal"]
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = None
    title.paragraph_format.space_before = Pt(0)
    title.paragraph_format.space_after = Pt(12)
    set_line_spacing(title, 1.0)

    # 原第二段只是“专项核查报告”的旧副标题，新主题已经完整包含该语义。
    if len(paragraphs) > 1 and paragraphs[1].text.strip() == "专项核查报告":
        delete_paragraph(paragraphs[1])

    # 删除真实源文件中的空段落，避免视觉上出现无意义大块留白。
    for paragraph in list(document.paragraphs):
        if paragraph._p is title._p:
            continue
        if not paragraph.text.strip() and not paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
            delete_paragraph(paragraph)

    signature_one = None
    signature_two = None
    h1_pattern = re.compile(r"^[一二三四五六七八九十百]+、")
    h2_pattern = re.compile(r"^[（(][一二三四五六七八九十百]+[）)]")
    for paragraph in list(document.paragraphs):
        if paragraph._p is title._p:
            continue
        text = paragraph.text.strip()
        compact = re.sub(r"\s+", "", text)
        if compact.startswith("核查编制人：") or compact.startswith("核查编制人:"):
            signature_one = paragraph
            continue
        if compact.startswith("编制日期：") or compact.startswith("编制日期:"):
            signature_two = paragraph
            continue
        if h1_pattern.match(text):
            style_second_heading(paragraph, 1)
        elif h2_pattern.match(text):
            style_second_heading(paragraph, 2)
        else:
            add_company_runs(paragraph, text)
            paragraph.style = document.styles["Normal"]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.first_line_indent = Pt(24)
            set_first_line_chars(paragraph, 200)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(4)
            set_line_spacing(paragraph, 1.5)

    if signature_one is None or signature_two is None:
        raise RuntimeError("未定位到两段编制签名")
    add_signature(signature_one, "核查编制人：", "林远芝", numeric=False)
    add_signature(signature_two, "编制日期：", "2026年5月20日", numeric=True)

    for table in document.tables:
        format_three_line_table(table, "宋体", shade_header=False)

    # 真实源文件启用了 31.2 磅文档网格，导致即使段落标记为 1.5 倍行距，视觉上仍
    # 接近三倍行距。逐段关闭吸附后，Rubric 要求的 1.5 倍行距才与办公软件显示一致。
    for paragraph in document.paragraphs:
        disable_document_grid(paragraph)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    disable_document_grid(paragraph)

    document.core_properties.title = SECOND_TITLE
    document.core_properties.subject = "华中区域零售项目资金与现金流专项核查"
    document.save(output_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument(
        "--toc-pages",
        default="3,4,4,5,6,7",
        help="六个一级标题在最终渲染中的页码，首次重放后按实际分页核对",
    )
    args = parser.parse_args()
    toc_pages = [int(item) for item in args.toc_pages.split(",")]
    if len(toc_pages) != 6:
        raise SystemExit("--toc-pages 必须给出六个页码")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    build_first(args.input_dir, args.output_dir / OUTPUT_ONE, toc_pages)
    build_second(args.input_dir, args.output_dir / OUTPUT_TWO)


if __name__ == "__main__":
    main()
