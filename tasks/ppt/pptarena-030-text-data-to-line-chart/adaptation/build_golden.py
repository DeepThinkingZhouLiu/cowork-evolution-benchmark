#!/usr/bin/env python3
"""PPTArena Case 30 的单题 Golden 构造器。

这个脚本只处理本题的第 8 页：它从冻结输入里读取明文表格，
删除该文本框，再生成原生可编辑的带数据标记折线图和右侧自定义图例。
不使用上游 Ground Truth 中输入未披露的隐藏小数。
"""

from __future__ import annotations

import argparse
import hashlib
import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_AXIS_CROSSES,
    XL_CHART_TYPE,
    XL_MARKER_STYLE,
    XL_TICK_LABEL_POSITION,
    XL_TICK_MARK,
)
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt


INPUT_NAME = "LineGraph_TestA.pptx"
OUTPUT_NAME = "LineGraph_GroundTruthA.pptx"
INPUT_SHA256 = "72fb98cc839fc54dc4b07dcdb4bf7eb8c6552ac6549d00c0010e3ab021c1f3e6"
TARGET_SLIDE_INDEX = 7
SOURCE_TABLE_SHAPE_ID = 18
YEARS = ["2013", "2014", "2015", "2016", "2017"]
EXPECTED = {
    "NHOPI": [0.18, 0.14, 0.10, 0.11, 0.11],
    "Asian": [0.16, 0.11, 0.08, 0.07, 0.07],
    "White": [0.12, 0.10, 0.08, 0.07, 0.07],
}


@dataclass(frozen=True)
class Palette:
    nhopi: str = "174A6E"
    asian: str = "0E83C9"
    white: str = "A9A9A9"

    def as_mapping(self) -> dict[str, str]:
        return {"NHOPI": self.nhopi, "Asian": self.asian, "White": self.white}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def normalized(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def parse_source_table(presentation: Presentation) -> dict[str, list[float]]:
    """从第 8 页明文表格重新读取数据，防止 Golden 脱离输入。"""
    slide = presentation.slides[TARGET_SLIDE_INDEX]
    matches = [shape for shape in slide.shapes if shape.shape_id == SOURCE_TABLE_SHAPE_ID]
    if len(matches) != 1 or not matches[0].has_text_frame:
        raise ValueError("无法唯一定位第 8 页的源数据文本框")
    lines = [normalized(line) for line in matches[0].text.splitlines() if normalized(line)]
    if len(lines) != 6:
        raise ValueError(f"源数据表应为 1 行表头+5 行数据，实际为 {len(lines)} 行")
    rows: list[list[int]] = []
    for line in lines[1:]:
        values = [int(value) for value in re.findall(r"\d+", line)]
        if len(values) != 4:
            raise ValueError(f"无法解析数据行：{line!r}")
        rows.append(values)
    if [str(row[0]) for row in rows] != YEARS:
        raise ValueError("源数据年份不是 2013-2017")
    parsed = {
        "NHOPI": [row[1] / 100 for row in rows],
        "Asian": [row[2] / 100 for row in rows],
        "White": [row[3] / 100 for row in rows],
    }
    if parsed != EXPECTED:
        raise ValueError(f"源数据与冻结任务契约不一致：{parsed}")
    return parsed


def remove_source_table(presentation: Presentation) -> None:
    slide = presentation.slides[TARGET_SLIDE_INDEX]
    matches = [shape for shape in slide.shapes if shape.shape_id == SOURCE_TABLE_SHAPE_ID]
    if len(matches) != 1:
        raise ValueError("源数据文本框数量异常")
    matches[0]._element.getparent().remove(matches[0]._element)


def set_chart_font(chart, size_pt: float = 10.5) -> None:
    """在 python-pptx 暂未覆盖的图表格式上做少量明确设置。"""
    chart.font.name = "Arial"
    chart.font.size = Pt(size_pt)
    chart.font.color.rgb = RGBColor.from_string("404040")


def add_line_chart(
    presentation: Presentation,
    values: dict[str, list[float]],
    palette: Palette,
    *,
    left: float = 0.72,
    top: float = 1.88,
    width: float = 7.15,
    height: float = 4.62,
) -> None:
    slide = presentation.slides[TARGET_SLIDE_INDEX]
    data = ChartData()
    data.categories = YEARS
    for name in ("NHOPI", "Asian", "White"):
        data.add_series(name, values[name], number_format="0%")

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
        data,
    )
    chart_shape.name = "Uninsured rates 2013-2017 - editable line chart"
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False
    chart.chart_style = 10
    set_chart_font(chart)

    category_axis = chart.category_axis
    category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    category_axis.major_tick_mark = XL_TICK_MARK.NONE
    category_axis.minor_tick_mark = XL_TICK_MARK.NONE
    category_axis.format.line.color.rgb = RGBColor.from_string("B7B7B7")
    category_axis.tick_labels.font.name = "Arial"
    category_axis.tick_labels.font.size = Pt(11)
    category_axis.tick_labels.font.color.rgb = RGBColor.from_string("333333")

    value_axis = chart.value_axis
    value_axis.minimum_scale = 0.0
    value_axis.maximum_scale = 0.20
    value_axis.major_unit = 0.05
    value_axis.crosses = XL_AXIS_CROSSES.MINIMUM
    value_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    value_axis.major_tick_mark = XL_TICK_MARK.NONE
    value_axis.minor_tick_mark = XL_TICK_MARK.NONE
    value_axis.tick_labels.number_format = "0%"
    value_axis.tick_labels.font.name = "Arial"
    value_axis.tick_labels.font.size = Pt(10)
    value_axis.tick_labels.font.color.rgb = RGBColor.from_string("595959")
    value_axis.format.line.fill.background()
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = RGBColor.from_string("D9D9D9")
    value_axis.major_gridlines.format.line.width = Pt(0.75)

    colors = palette.as_mapping()
    for series in chart.series:
        color = RGBColor.from_string(colors[series.name])
        series.format.line.color.rgb = color
        series.format.line.width = Pt(2.25)
        series.marker.style = XL_MARKER_STYLE.CIRCLE
        series.marker.size = 7
        series.marker.format.fill.solid()
        series.marker.format.fill.fore_color.rgb = color
        series.marker.format.line.color.rgb = color


def add_custom_legend(
    presentation: Presentation,
    palette: Palette,
    *,
    left: float = 8.12,
    top: float = 3.03,
    row_gap: float = 0.57,
) -> None:
    """用原生线段、标记和文本构造右侧图例，不依赖自动图例。"""
    slide = presentation.slides[TARGET_SLIDE_INDEX]
    colors = palette.as_mapping()
    for index, name in enumerate(("NHOPI", "Asian", "White")):
        y = top + index * row_gap
        color = RGBColor.from_string(colors[name])
        line = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            Inches(left),
            Inches(y + 0.14),
            Inches(left + 0.35),
            Inches(y + 0.14),
        )
        line.name = f"Custom legend line - {name}"
        line.line.color.rgb = color
        line.line.width = Pt(2.25)

        marker = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(left + 0.135),
            Inches(y + 0.075),
            Inches(0.13),
            Inches(0.13),
        )
        marker.name = f"Custom legend marker - {name}"
        marker.fill.solid()
        marker.fill.fore_color.rgb = color
        marker.line.color.rgb = color

        label = slide.shapes.add_textbox(
            Inches(left + 0.43),
            Inches(y),
            Inches(1.23),
            Inches(0.32),
        )
        label.name = f"Custom legend label - {name}"
        label.text_frame.clear()
        label.text_frame.margin_left = 0
        label.text_frame.margin_right = 0
        label.text_frame.margin_top = 0
        label.text_frame.margin_bottom = 0
        label.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        run = label.text_frame.paragraphs[0].add_run()
        run.text = name
        run.font.name = "Arial"
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor.from_string(colors[name])


def build(source: Path, output: Path, palette: Palette = Palette()) -> None:
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("冻结输入缺失或 SHA256 不匹配")
    presentation = Presentation(source)
    if len(presentation.slides) != 9:
        raise ValueError("冻结输入应为 9 页")
    values = parse_source_table(presentation)
    remove_source_table(presentation)
    add_line_chart(presentation, values, palette)
    add_custom_legend(presentation, palette)
    output.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    build(args.source, args.output)
    print(args.output)


if __name__ == "__main__":
    main()
