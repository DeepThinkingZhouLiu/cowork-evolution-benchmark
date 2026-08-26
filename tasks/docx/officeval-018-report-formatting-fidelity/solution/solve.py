#!/usr/bin/env python3
"""逐项整理 OfficeVal officeval_018 的真实报告文档。

这不是批量生成器。下方的段落索引、标题层级、图题和表题都是针对
officeval_018 的两份真实输入逐项检查后固定下来的。
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph


TITLE = "青禾公共书房网络化运营与更新规划书"
OUTPUT_NAME = "公共书房运营规划书_排版完成版.docx"

# 原文段落索引 -> 目标层级。索引来自对原始 DOCX 的人工逐段检查。
HEADING_LEVELS = {
    1: [43, 155, 219, 275, 315, 329, 379],
    2: [
        44, 59, 63, 65, 74, 85, 86,
        157, 161, 165, 169, 174, 178, 199,
        220, 243,
        279, 284, 293, 296, 298, 301, 309,
        316, 322, 327,
        331, 333, 352,
    ],
    3: [
        45, 61, 66, 76, 78, 87, 119,
        179, 186, 204, 212,
        221, 229, 234, 238, 244, 262,
        285, 287, 289, 291,
        323, 325,
        334, 342, 347, 357, 362, 369, 374,
    ],
    4: [
        47, 51, 55, 80, 81, 82, 83, 84, 89, 91, 93, 99, 109, 115,
        120, 124, 137,
        180, 183, 191, 193, 195,
        248, 252, 257, 263, 267, 271,
    ],
    5: [95, 97, 102, 104, 106, 111, 113, 117, 118, 125, 140, 143, 150],
}

# 原文中有 5 个“标题+正文”粘在同一段的位置，需先分开再排版。
SPLIT_SPECS = {
    81: ("。", "读者访谈与社群共创"),
    83: ("。", "服务质量评估"),
    85: ("。", "规划创新点"),
    117: ("：", "空间类型划分"),
    118: ("：", "样本点选择说明"),
}

# 31 个可见图形组的题名。原段 189 的两张图是一幅并排组图；
# 原段 375 的两张图是两幅独立图，改造时拆成两段并分别编号。
# 原段 280 引用的 image9.png 是 RGBA 全透明占位图，人眼不可见，不当作图片交付。
FIGURE_TITLES = [
    "公共书房活动流程图",
    "读者年龄分布图",
    "公共书房空间动线图",
    "读者使用年限分布图",
    "读者职业类型分布图",
    "读者日均停留时长分布图",
    "读者周边居住距离分布图",
    "活动流程与阅读服务概览图",
    "覆盖度分析空间动线图",
    "覆盖度分析阅读服务看板",
    "可达性分析共创反馈闭环图",
    "社区服务使用频次分布图",
    "社区服务需求分布图",
    "未使用社区服务的原因分布图",
    "社区服务点覆盖度评价图",
    "社区服务可达性评价图",
    "读者参与社区共建认同度图",
    "读者参与社区共建意愿图",
    "读者社区服务参与经历图",
    "未参与社区服务原因图",
    "综合服务点选址偏好图",
    "社区共建积分入口偏好图",
    "社区共建积分兑换福利偏好图",
    "社区共建参与渠道偏好图",
    "阅读服务看板",
    "公共书房客流趋势图",
    "公共书房资源配置图",
    "公共书房共创反馈闭环图",
    "公共书房空间动线示意图",
    "公共书房客流趋势图（落地计划）",
    "公共书房阅读服务看板（落地计划）",
]

TABLE_TITLES = [
    "公共书房运营指标表",
    "服务模块成熟度统计表",
    "点位服务执行要点表",
    "运营模型服务动作说明表",
    "模型校准服务动作说明表",
    "运营模型拟合指标表",
    "活动资源配置对照表",
    "模型结论拟合指标表",
    "服务需求与空间供给执行表",
    "读者年龄组模型拟合指标表",
    "访问类型点位执行要点表",
    "访问类型运营指标统计表",
    "访问类型差异检验表",
    "停留时长点位执行要点表",
    "停留时长运营指标统计表",
    "居住距离运营指标统计表",
    "居住距离差异检验表",
    "居住距离分组运营指标表",
    "读者特征点位执行要点表",
    "读者特征运营指标统计表",
    "年龄分组点位执行要点表",
    "年龄分组运营指标统计表",
    "访问类型空间模块偏好差异检验表",
    "停留时长阅读配套偏好差异检验表",
    "居住距离夜间开放偏好差异检验表",
]

CHINESE_NUMERALS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]


def chinese_number(value: int) -> str:
    if 0 <= value <= 10:
        return CHINESE_NUMERALS[value]
    if value < 20:
        return "十" + CHINESE_NUMERALS[value - 10]
    tens, units = divmod(value, 10)
    return CHINESE_NUMERALS[tens] + "十" + (CHINESE_NUMERALS[units] if units else "")


def get_or_add(parent, tag: str):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    return child


def set_run_font(run, name: str, size_pt: float, *, bold: bool | None = None) -> None:
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    fonts = get_or_add(rpr, "w:rFonts")
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attribute), name)


def set_spacing(
    paragraph: Paragraph,
    *,
    line: int,
    before_lines: int = 0,
    after_twips: int = 0,
    first_line_chars: int = 0,
) -> None:
    ppr = paragraph._element.get_or_add_pPr()
    spacing = get_or_add(ppr, "w:spacing")
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), "auto")
    spacing.set(qn("w:beforeLines"), str(before_lines))
    spacing.set(qn("w:after"), str(after_twips))
    for attr in ("w:before", "w:afterLines"):
        spacing.attrib.pop(qn(attr), None)
    ind = get_or_add(ppr, "w:ind")
    ind.set(qn("w:firstLineChars"), str(first_line_chars))
    for attr in ("w:firstLine", "w:hanging", "w:hangingChars"):
        ind.attrib.pop(qn(attr), None)


def remove_auto_numbering(paragraph: Paragraph) -> None:
    """用 numId=0 显式禁用样式继承的自动编号。

    只删除段落自身的 numPr 不够：Heading 1–5 样式在原文中还关联了
    numbering.xml，LibreOffice 会再次显示一套编号。
    """
    ppr = paragraph._element.get_or_add_pPr()
    num_pr = get_or_add(ppr, "w:numPr")
    for child in list(num_pr):
        num_pr.remove(child)
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    num_id = OxmlElement("w:numId")
    num_id.set(qn("w:val"), "0")
    num_pr.extend((ilvl, num_id))


def replace_text(paragraph: Paragraph, text: str):
    paragraph.clear()
    return paragraph.add_run(text)


def insert_paragraph_after(paragraph: Paragraph, text: str = "") -> Paragraph:
    element = OxmlElement("w:p")
    paragraph._p.addnext(element)
    created = Paragraph(element, paragraph._parent)
    if text:
        created.add_run(text)
    return created


def insert_paragraph_before(paragraph: Paragraph, text: str = "") -> Paragraph:
    element = OxmlElement("w:p")
    paragraph._p.addprevious(element)
    created = Paragraph(element, paragraph._parent)
    if text:
        created.add_run(text)
    return created


def insert_caption_before_table(table, text: str) -> Paragraph:
    element = OxmlElement("w:p")
    table._tbl.addprevious(element)
    paragraph = Paragraph(element, table._parent)
    paragraph.add_run(text)
    return paragraph


def delete_paragraph(paragraph: Paragraph) -> None:
    parent = paragraph._element.getparent()
    if parent is not None:
        parent.remove(paragraph._element)


def clear_text_preserve_drawings(paragraph: Paragraph) -> None:
    """删掉与图形锚在同一段的错位文字，保留 drawing 和关系部件。"""
    for run_element in list(paragraph._element.findall(qn("w:r"))):
        has_drawing = bool(run_element.findall(".//" + qn("w:drawing")))
        if has_drawing:
            for tag in ("w:t", "w:instrText", "w:tab", "w:br", "w:fldChar"):
                for node in list(run_element.findall(".//" + qn(tag))):
                    node.getparent().remove(node)
        else:
            paragraph._element.remove(run_element)
    for hyperlink in list(paragraph._element.findall(qn("w:hyperlink"))):
        paragraph._element.remove(hyperlink)


def convert_anchors_to_inline(paragraph: Paragraph) -> None:
    """把指定段中的浮动图转成嵌入型图，让图片真正占据正文高度。"""
    for anchor in list(paragraph._element.findall(".//" + qn("wp:anchor"))):
        inline = OxmlElement("wp:inline")
        for attribute in ("distT", "distB", "distL", "distR"):
            inline.set(attribute, "0")
        for tag in ("wp:extent", "wp:effectExtent", "wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
            child = anchor.find(qn(tag))
            if child is not None:
                anchor.remove(child)
                inline.append(child)
        anchor.getparent().replace(anchor, inline)


def split_second_drawing_to_new_paragraph(paragraph: Paragraph) -> Paragraph:
    """把同段中的第二个独立图移到新段，使两幅图可以各自紧跟图注。"""
    drawing_runs = [
        run for run in paragraph._element.findall(qn("w:r"))
        if run.findall(".//" + qn("w:drawing"))
    ]
    if len(drawing_runs) != 2:
        raise ValueError(f"预期组图段有 2 个绘图 run，实际 {len(drawing_runs)}")
    created = insert_paragraph_after(paragraph)
    paragraph._element.remove(drawing_runs[1])
    created._element.append(drawing_runs[1])
    return created


def remove_redundant_section_break(paragraph: Paragraph) -> None:
    """移除图片段中纸张和页边距完全一样的冗余分节，避免图注被强制换页。"""
    ppr = paragraph._element.get_or_add_pPr()
    section_properties = ppr.find(qn("w:sectPr"))
    if section_properties is not None:
        ppr.remove(section_properties)


def has_drawing(paragraph: Paragraph) -> bool:
    return bool(
        paragraph._element.findall(".//" + qn("w:drawing"))
        or paragraph._element.findall(".//" + qn("w:pict"))
    )


def strip_old_numbering(text: str) -> str:
    value = text.strip().strip("/").strip()
    patterns = [
        r"^第[一二三四五六七八九十]+章[\s、，,]*",
        r"^[一二三四五六七八九十]+、\s*",
        r"^（[一二三四五六七八九十]+）\s*",
        r"^（\d+）\s*",
        r"^\d+[\.\uff0e、）)]\s*",
    ]
    for pattern in patterns:
        value = re.sub(pattern, "", value, count=1)
    return value.strip().rstrip("。")


def format_title(paragraph: Paragraph) -> None:
    paragraph.style = "Normal"
    run = replace_text(paragraph, TITLE)
    set_run_font(run, "宋体", 16, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(paragraph, line=240)
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.keep_together = True
    remove_auto_numbering(paragraph)


def format_abstract(paragraph: Paragraph) -> None:
    paragraph.style = "Normal"
    run = replace_text(paragraph, "摘要")
    set_run_font(run, "黑体", 14, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(paragraph, line=240)
    paragraph.paragraph_format.page_break_before = True
    paragraph.paragraph_format.keep_with_next = False
    remove_auto_numbering(paragraph)


def format_heading(paragraph: Paragraph, level: int, text: str) -> None:
    style_name = f"Heading {level}"
    try:
        paragraph.style = style_name
    except KeyError:
        style = paragraph.part.document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        paragraph.style = style
    run = replace_text(paragraph, text)
    if level == 1:
        set_run_font(run, "黑体", 15, bold=True)
    elif level == 2:
        set_run_font(run, "黑体", 14, bold=True)
    else:
        set_run_font(run, "宋体", 12, bold=False)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_spacing(paragraph, line=240, before_lines=50 if level <= 2 else 25)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    paragraph.paragraph_format.page_break_before = False
    remove_auto_numbering(paragraph)


def format_caption(paragraph: Paragraph) -> None:
    try:
        paragraph.style = "Caption"
    except KeyError:
        pass
    for run in paragraph.runs:
        if run.text:
            set_run_font(run, "宋体", 12, bold=False)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(paragraph, line=240)
    paragraph.paragraph_format.keep_together = True
    remove_auto_numbering(paragraph)


def format_body(paragraph: Paragraph) -> None:
    for run in paragraph.runs:
        if run.text.strip():
            set_run_font(run, "宋体", 12)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    set_spacing(paragraph, line=300, before_lines=50, first_line_chars=200)
    paragraph.paragraph_format.keep_together = False
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.page_break_before = False
    remove_auto_numbering(paragraph)


def format_empty(paragraph: Paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_spacing(paragraph, line=240)
    paragraph.paragraph_format.keep_with_next = False
    paragraph.paragraph_format.keep_together = False
    remove_auto_numbering(paragraph)


def format_figure_paragraph(paragraph: Paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(paragraph, line=240)
    paragraph.paragraph_format.keep_with_next = True
    paragraph.paragraph_format.keep_together = True
    remove_auto_numbering(paragraph)


def format_tables(document: Document) -> None:
    for table in document.tables:
        for row_index, row in enumerate(table.rows):
            tr_pr = row._tr.get_or_add_trPr()
            if tr_pr.find(qn("w:cantSplit")) is None:
                tr_pr.append(OxmlElement("w:cantSplit"))
            if row_index == 0 and tr_pr.find(qn("w:tblHeader")) is None:
                tr_pr.append(OxmlElement("w:tblHeader"))
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            set_run_font(run, "宋体", 12)
                            if row_index == 0:
                                run.bold = True
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if row_index == 0 else WD_ALIGN_PARAGRAPH.LEFT
                    set_spacing(paragraph, line=240)
                    paragraph.paragraph_format.keep_together = True
                    # 表格末行外的内容保持与下一行相连，减少表格在页间被切开。
                    paragraph.paragraph_format.keep_with_next = row_index < len(table.rows) - 1
                    remove_auto_numbering(paragraph)


def heading_prefix(level: int, number: int) -> str:
    if level == 1:
        return f"{chinese_number(number)}、"
    if level == 2:
        return f"（{chinese_number(number)}）"
    if level == 3:
        return f"{number}."
    if level == 4:
        return f"（{number}）"
    return f"{number}）"


def solve(input_path: Path, output_path: Path) -> None:
    document = Document(input_path)
    original_paragraphs = list(document.paragraphs)
    if len(original_paragraphs) != 386:
        raise ValueError(f"预期原文有 386 个段落，实际 {len(original_paragraphs)}")
    if len(document.tables) != 25:
        raise ValueError(f"预期原文有 25 张表，实际 {len(document.tables)}")

    # 两份原始文档的页面要求均以输入中的版式说明为准。
    for section in document.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.18)
        section.right_margin = Cm(3.18)

    format_title(original_paragraphs[0])
    format_abstract(original_paragraphs[22])
    # 删掉题目与摘要之间用回车堆出来的空白，改用明确分页。
    for paragraph in original_paragraphs[1:22] + original_paragraphs[23:43]:
        if not paragraph.text.strip() and not has_drawing(paragraph):
            delete_paragraph(paragraph)
    original_paragraphs[43].paragraph_format.page_break_before = True

    # 先拆开粘连的标题和正文。
    split_bodies: list[Paragraph] = []
    heading_titles: dict[int, str] = {}
    for index, (delimiter, fixed_title) in SPLIT_SPECS.items():
        paragraph = original_paragraphs[index]
        source_text = paragraph.text.strip()
        if delimiter not in source_text:
            raise ValueError(f"段落 {index} 缺少预期分隔符 {delimiter}")
        body_text = source_text.split(delimiter, 1)[1].strip()
        heading_titles[index] = fixed_title
        if body_text:
            split_bodies.append(insert_paragraph_after(paragraph, body_text))

    # 原段 161 把小标题和图片锚在了同一段：标题移到图片之前，图片保留。
    image_heading = insert_paragraph_before(original_paragraphs[161], "协作机构行动")
    clear_text_preserve_drawings(original_paragraphs[161])
    # 原段 217 把图注写进图片段，清掉错位文字，后面重建正常图注。
    clear_text_preserve_drawings(original_paragraphs[217])
    # 三处浮动图原本不占正文流高度，造成图注跑到图前、与下一图相叠。
    for index in (161, 217, 375):
        convert_anchors_to_inline(original_paragraphs[index])
    split_second_drawing_to_new_paragraph(original_paragraphs[375])
    remove_redundant_section_break(original_paragraphs[375])
    # image9.png 的 RGBA 四个通道全部为 0，是不可见占位对象，不生成虚假图注。
    delete_paragraph(original_paragraphs[280])

    # 删掉不在正确位置或内容与图表不符的旧图注/表注。
    caption_pattern = re.compile(r"^[图表]\s*\d+")
    for paragraph in list(document.paragraphs):
        if not has_drawing(paragraph) and caption_pattern.match(paragraph.text.strip()):
            delete_paragraph(paragraph)

    # 按原文顺序收集 31 个可见图形组，逐个插入真实内容对应的图注。
    figure_paragraphs = [paragraph for paragraph in document.paragraphs if has_drawing(paragraph)]
    if len(figure_paragraphs) != len(FIGURE_TITLES):
        raise ValueError(
            f"预期 {len(FIGURE_TITLES)} 个图形段落，实际 {len(figure_paragraphs)}"
        )
    figure_captions: list[Paragraph] = []
    for number, (paragraph, title) in enumerate(zip(figure_paragraphs, FIGURE_TITLES), 1):
        caption = insert_paragraph_after(paragraph, f"图 {number}  {title}")
        format_caption(caption)
        figure_captions.append(caption)

    if len(TABLE_TITLES) != len(document.tables):
        raise ValueError("表题数量与表格数量不一致")
    table_captions: list[Paragraph] = []
    for number, (table, title) in enumerate(zip(document.tables, TABLE_TITLES), 1):
        caption = insert_caption_before_table(table, f"表 {number}  {title}")
        format_caption(caption)
        caption.paragraph_format.keep_with_next = True
        table_captions.append(caption)

    # 建立所有标题引用。第 161 段的标题使用刚刚拆出的新段落。
    heading_refs: dict[int, tuple[int, Paragraph, str]] = {}
    for level, indexes in HEADING_LEVELS.items():
        for index in indexes:
            if index == 161:
                paragraph = image_heading
                title = "协作机构行动"
            else:
                paragraph = original_paragraphs[index]
                title = heading_titles.get(index, strip_old_numbering(paragraph.text))
            heading_refs[index] = (level, paragraph, title)

    counters = [0, 0, 0, 0, 0]
    heading_paragraph_ids: set[int] = set()
    for index in sorted(heading_refs):
        level, paragraph, title = heading_refs[index]
        counters[level - 1] += 1
        for deeper in range(level, 5):
            counters[deeper] = 0
        rendered = heading_prefix(level, counters[level - 1]) + title
        format_heading(paragraph, level, rendered)
        heading_paragraph_ids.add(id(paragraph._element))
    # 摘要后的第一章使用显式分页，保持题目页/摘要页/正文的层次。
    original_paragraphs[43].paragraph_format.page_break_before = True

    caption_ids = {id(p._element) for p in figure_captions + table_captions}
    title_id = id(original_paragraphs[0]._element)
    abstract_id = id(original_paragraphs[22]._element)

    # 最后统一正文、空段和图形段。先前设定好的标题/题注不再被覆盖。
    for paragraph in document.paragraphs:
        pid = id(paragraph._element)
        if pid in (title_id, abstract_id) or pid in heading_paragraph_ids or pid in caption_ids:
            continue
        if has_drawing(paragraph):
            format_figure_paragraph(paragraph)
        elif paragraph.text.strip():
            format_body(paragraph)
        else:
            format_empty(paragraph)

    for body_paragraph in split_bodies:
        format_body(body_paragraph)
    format_tables(document)

    # 应用到文档默认样式，避免新增空 run 回落到非中文字体。
    normal = document.styles["Normal"]
    normal.font.name = "宋体"
    normal.font.size = Pt(12)
    normal_rpr = normal.element.get_or_add_rPr()
    normal_fonts = get_or_add(normal_rpr, "w:rFonts")
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        normal_fonts.set(qn(attribute), "宋体")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    solve(args.input, args.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
