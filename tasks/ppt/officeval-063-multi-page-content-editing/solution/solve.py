#!/usr/bin/env python3
"""officeval_063 的逐题参考解法。

本文件只修改真实输入“直播运营阶段复盘.pptx”中被 4 张标注图
指定的页面，并把 Query 里的 5 个月真实数值写入第 10 页原生图表。
它不生成新题，也不改写未被用户指定的页面。
"""

from __future__ import annotations

import argparse
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


SOURCE_NAME = "直播运营阶段复盘.pptx"
OUTPUT_NAME = "直播运营阶段复盘_修订版.pptx"
FONT_CN = "Microsoft YaHei"
DEEP_TEAL = "0B5D66"
TEAL = "20A6A8"
ORANGE = "F0A45D"
DARK = "17343C"
GRAY = "5E737A"
LIGHT_BLUE = "F3FAFB"


def find_name(slide, name: str):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise ValueError(f"第 {slide.slide_id} 页缺少对象 {name}")


def remove_shape(shape) -> None:
    parent = shape._element.getparent()
    if parent is not None:
        parent.remove(shape._element)


def set_background(slide, color: str) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(color)


def style_textbox(shape, *, size: float, bold: bool, color: str, alignment=None) -> None:
    if not getattr(shape, "has_text_frame", False):
        raise ValueError(f"{shape.name} 不是文本对象")
    for paragraph in shape.text_frame.paragraphs:
        if alignment is not None:
            paragraph.alignment = alignment
        for run in paragraph.runs:
            run.font.name = FONT_CN
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = RGBColor.from_string(color)


def add_shift_title(presentation: Presentation) -> None:
    """按标注图 6.JPG，在第 2 页右侧第三张岗位卡的留白处补标题。"""
    slide = presentation.slides[1]
    title = slide.shapes.add_textbox(Inches(8.68), Inches(1.67), Inches(2.75), Inches(0.52))
    title.name = "Added Shift Role Title"
    title.text_frame.clear()
    title.text_frame.margin_left = 0
    title.text_frame.margin_right = 0
    title.text_frame.margin_top = 0
    title.text_frame.margin_bottom = 0
    title.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = title.text_frame.paragraphs[0]
    paragraph.add_run().text = "排班岗"
    style_textbox(title, size=24, bold=True, color=DEEP_TEAL)


def straighten_results_page(presentation: Presentation) -> None:
    """按标注图 7.JPG，校正第 5 页全部卡片、文字和右侧照片。"""
    slide = presentation.slides[4]
    set_background(slide, LIGHT_BLUE)
    for shape in slide.shapes:
        shape.rotation = 0


def remove_duplicate_action(presentation: Presentation) -> None:
    """按标注图 8.JPG，仅删除第 6 页优化动作中第二个重复段落。"""
    slide = presentation.slides[5]
    body = find_name(slide, "Text 8")
    target = "将敏感表达改为场景化、说明式表达"
    seen = 0
    for paragraph in list(body.text_frame.paragraphs):
        if target not in paragraph.text:
            continue
        seen += 1
        if seen > 1:
            paragraph._p.getparent().remove(paragraph._p)
    remaining = sum(target in paragraph.text for paragraph in body.text_frame.paragraphs)
    if remaining != 1:
        raise ValueError(f"第 6 页去重后目标语句应剩 1 次，实际 {remaining} 次")


def set_three_lines(shape, lines: list[str]) -> None:
    text_frame = shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_left = 0
    text_frame.margin_right = 0
    for index, text in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.add_run().text = text
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_after = Pt(9)
        for run in paragraph.runs:
            run.font.name = FONT_CN
            run.font.size = Pt(18)
            run.font.bold = False
            run.font.color.rgb = RGBColor.from_string(DARK)


def simplify_planning_page(presentation: Presentation) -> None:
    """按标注图 9.JPG，删除中期板块和底部临时文字，重排近期/长期。"""
    slide = presentation.slides[8]
    for name in ("Shape 6", "Text 7", "Text 8", "Text 12", "Shape 14"):
        remove_shape(find_name(slide, name))

    recent_card = find_name(slide, "Shape 3")
    recent_title = find_name(slide, "Text 4")
    recent_body = find_name(slide, "Text 5")
    long_card = find_name(slide, "Shape 9")
    long_title = find_name(slide, "Text 10")
    long_body = find_name(slide, "Text 11")

    for shape in (recent_card, long_card):
        shape.top = Inches(2.05)
        shape.width = Inches(3.7)
        shape.height = Inches(2.75)
    recent_card.left = Inches(1.90)
    long_card.left = Inches(7.73)

    for shape in (recent_title, long_title):
        shape.top = Inches(2.30)
        shape.width = Inches(3.2)
        shape.height = Inches(0.35)
    recent_title.left = Inches(2.15)
    long_title.left = Inches(7.98)

    for shape in (recent_body, long_body):
        shape.top = Inches(2.82)
        shape.width = Inches(3.2)
        shape.height = Inches(1.55)
    recent_body.left = Inches(2.15)
    long_body.left = Inches(7.98)

    set_three_lines(
        long_body,
        ["搭建标准化复盘范式", "强化岗位协同配合效能", "赋能团队稳步长效发展"],
    )


def crop_picture_to_slide(picture, presentation: Presentation) -> None:
    """把结束页原照片放大为全页背景，保留图片而不与图表争抢空间。"""
    source_width, source_height = picture.image.size
    source_ratio = source_width / source_height
    target_ratio = presentation.slide_width / presentation.slide_height
    picture.left = 0
    picture.top = 0
    picture.width = presentation.slide_width
    picture.height = presentation.slide_height
    if source_ratio < target_ratio:
        visible_height = source_width / target_ratio
        crop = max(0.0, min(0.49, (source_height - visible_height) / source_height / 2))
        picture.crop_left = 0
        picture.crop_right = 0
        picture.crop_top = crop
        picture.crop_bottom = crop
    else:
        visible_width = source_height * target_ratio
        crop = max(0.0, min(0.49, (source_width - visible_width) / source_width / 2))
        picture.crop_top = 0
        picture.crop_bottom = 0
        picture.crop_left = crop
        picture.crop_right = crop


def set_child_value(parent, tag: str, value: str) -> None:
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    child.set("val", value)


def enable_secondary_axis(chart) -> None:
    """把“盈利比值”系列移到右侧次坐标轴，保持原生可编辑图表。"""
    plot_area = chart._chartSpace.find(".//c:plotArea", {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"})
    line_chart = plot_area.find(qn("c:lineChart"))
    series = line_chart.findall(qn("c:ser"))
    if len(series) != 2:
        raise ValueError(f"次坐标轴改造前应有 2 个系列，实际 {len(series)}")

    secondary_chart = deepcopy(line_chart)
    secondary_series = secondary_chart.findall(qn("c:ser"))
    secondary_chart.remove(secondary_series[0])
    line_chart.remove(series[1])

    cat_axis = plot_area.find(qn("c:catAx"))
    value_axis = plot_area.find(qn("c:valAx"))
    if cat_axis is None or value_axis is None:
        raise ValueError("图表缺少主分类轴或主数值轴")
    existing_ids = [
        int(node.get("val"))
        for node in plot_area.findall(".//c:axId", {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"})
        if node.get("val")
    ]
    secondary_cat_id = str(max(existing_ids) + 1001)
    secondary_val_id = str(max(existing_ids) + 2002)

    secondary_axis_ids = secondary_chart.findall(qn("c:axId"))
    secondary_axis_ids[0].set("val", secondary_cat_id)
    secondary_axis_ids[1].set("val", secondary_val_id)
    line_chart.addnext(secondary_chart)

    secondary_cat = deepcopy(cat_axis)
    set_child_value(secondary_cat, "c:axId", secondary_cat_id)
    set_child_value(secondary_cat, "c:axPos", "t")
    set_child_value(secondary_cat, "c:crossAx", secondary_val_id)
    set_child_value(secondary_cat, "c:delete", "1")
    set_child_value(secondary_cat, "c:crosses", "max")

    secondary_val = deepcopy(value_axis)
    set_child_value(secondary_val, "c:axId", secondary_val_id)
    set_child_value(secondary_val, "c:axPos", "r")
    set_child_value(secondary_val, "c:crossAx", secondary_cat_id)
    set_child_value(secondary_val, "c:crosses", "max")
    scaling = secondary_val.find(qn("c:scaling"))
    for tag, value in (("c:min", "1.3"), ("c:max", "1.9")):
        node = scaling.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            scaling.append(node)
        node.set("val", value)
    set_child_value(secondary_val, "c:majorUnit", "0.1")
    number_format = secondary_val.find(qn("c:numFmt"))
    if number_format is None:
        number_format = OxmlElement("c:numFmt")
        secondary_val.append(number_format)
    number_format.set("formatCode", "0.00")
    number_format.set("sourceLinked", "0")

    value_axis.addnext(secondary_cat)
    secondary_cat.addnext(secondary_val)


def style_chart_area(chart) -> None:
    """为图表增加高透明度白色底和浅青边框，避免数据线被照片干扰。"""
    chart_space = chart._chartSpace
    old = chart_space.find(qn("c:spPr"))
    if old is not None:
        chart_space.remove(old)
    shape_properties = OxmlElement("c:spPr")
    solid_fill = OxmlElement("a:solidFill")
    fill_color = OxmlElement("a:srgbClr")
    fill_color.set("val", "FFFFFF")
    alpha = OxmlElement("a:alpha")
    alpha.set("val", "93000")
    fill_color.append(alpha)
    solid_fill.append(fill_color)
    shape_properties.append(solid_fill)
    line = OxmlElement("a:ln")
    line.set("w", "12700")
    line_fill = OxmlElement("a:solidFill")
    line_color = OxmlElement("a:srgbClr")
    line_color.set("val", "B9DDE0")
    line_fill.append(line_color)
    line.append(line_fill)
    shape_properties.append(line)
    chart_space.chart.addnext(shape_properties)


def add_native_chart(presentation: Presentation) -> None:
    """将 Query 中的 5 个月数据写入第 10 页可编辑双轴折线图。"""
    slide = presentation.slides[9]
    crop_picture_to_slide(find_name(slide, "Image 0"), presentation)

    data = ChartData()
    data.categories = ["1月", "2月", "3月", "4月", "5月"]
    data.add_series("营收数值", (68000, 352000, 486000, 423000, 579000))
    data.add_series("盈利比值", (1.42, 1.56, 1.71, 1.63, 1.85))

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(7.12),
        Inches(1.75),
        Inches(5.80),
        Inches(3.65),
        data,
    )
    chart_shape.name = "Monthly Revenue and Profit Ratio Chart"
    chart = chart_shape.chart
    chart.chart_style = 13
    chart.has_title = True
    chart.chart_title.text_frame.text = "月度营收与盈利比值趋势"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.category_axis.has_major_gridlines = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.minimum_scale = 0
    chart.value_axis.maximum_scale = 600000
    chart.value_axis.major_unit = 100000
    chart.value_axis.tick_labels.number_format = "0"
    chart.value_axis.tick_labels.number_format_is_linked = False

    revenue, ratio = chart.series
    revenue.format.line.color.rgb = RGBColor.from_string(TEAL)
    revenue.format.line.width = Pt(2.5)
    revenue.marker.style = XL_MARKER_STYLE.CIRCLE
    revenue.marker.size = 8
    revenue.smooth = True
    ratio.format.line.color.rgb = RGBColor.from_string(ORANGE)
    ratio.format.line.width = Pt(2.5)
    ratio.marker.style = XL_MARKER_STYLE.DIAMOND
    ratio.marker.size = 8
    ratio.smooth = True

    title_frame = chart.chart_title.text_frame
    for paragraph in title_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.name = FONT_CN
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor.from_string(DARK)
    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.name = FONT_CN
        axis.tick_labels.font.size = Pt(9)
        axis.tick_labels.font.color.rgb = RGBColor.from_string(GRAY)
    chart.legend.font.name = FONT_CN
    chart.legend.font.size = Pt(9)

    enable_secondary_axis(chart)
    style_chart_area(chart)


def solve(input_dir: Path, output_dir: Path) -> Path:
    source_path = input_dir / SOURCE_NAME
    presentation = Presentation(source_path)
    if len(presentation.slides) != 10:
        raise ValueError(f"待改 PPT 应为 10 页，实际 {len(presentation.slides)} 页")

    add_shift_title(presentation)
    straighten_results_page(presentation)
    remove_duplicate_action(presentation)
    simplify_planning_page(presentation)
    add_native_chart(presentation)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_NAME
    presentation.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
