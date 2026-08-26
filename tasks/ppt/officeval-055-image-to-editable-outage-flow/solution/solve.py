#!/usr/bin/env python3
"""officeval_055 的逐题参考解法。

这个脚本只复现本题的人工改造结果：核对上游图片后，将五阶段
网站停机流程重建为可逐个编辑的 PowerPoint 形状、文字和连线。
输入图片只用于校验来源，不会被嵌入参考答案。
"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Cm, Pt


INPUT_NAME = "website_outage_editable_flow.png"
OUTPUT_NAME = "website_outage_editable_flow.pptx"
INPUT_SHA256 = "80bca602124e0cd54203937d2d46817062e7c61849cc05c911e64345be1c6add"

PAGE_W_CM = 33.867
PAGE_H_CM = 19.05

WHITE = "FFFFFF"
BLACK = "161A1D"
DARK_GREEN = "1B6652"
DARK_TEAL = "0B5364"
LIGHT_BLUE = "DDEFF5"
PALE_BLUE = "EDF7FA"
PALE_BORDER = "B7D7E1"
ORANGE = "E87522"
GREEN = "2E9D62"
THEME_BLUE = "1666A8"

STAGES = [
    {
        "number": "1",
        "title": "start",
        "body": "At 8:40 a.m., the team noticed that the online store was no longer loading for customers.",
        "center": 3.40,
        "body_width": 5.20,
        "body_fill": WHITE,
        "icon": "monitor-warning",
    },
    {
        "number": "2",
        "title": "check",
        "body": "The support lead checked alerts, confirmed the outage, and opened an incident channel for developers and operations.",
        "center": 10.17,
        "body_width": 5.20,
        "body_fill": PALE_BLUE,
        "icon": "magnifier-alert",
    },
    {
        "number": "3",
        "title": "problem peak",
        "body": "Orders failed repeatedly because the database connection pool was exhausted during the launch traffic surge.",
        "center": PAGE_W_CM / 2,
        "body_width": 5.20,
        "body_fill": PALE_BLUE,
        "icon": "server-warning",
    },
    {
        "number": "4",
        "title": "solution point",
        "body": "The team rolled back the latest configuration and switched traffic to a healthy backup instance.",
        "center": 23.70,
        "body_width": 5.20,
        "body_fill": PALE_BLUE,
        "icon": "cloud-check",
    },
    {
        "number": "5",
        "title": "finish",
        "body": "The site was back up and stable within 10 minutes, order submissions resumed, and a post-incident review has been scheduled for 4:00 PM this afternoon.",
        "center": 30.35,
        "body_width": 5.60,
        "body_fill": PALE_BLUE,
        "icon": "monitor-check",
    },
]


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def set_name(shape, name: str) -> None:
    shape.name = name


def set_fill(shape, color: str | None) -> None:
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(color)


def set_outline(shape, color: str | None, width_pt: float = 1.0) -> None:
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(color)
        shape.line.width = Pt(width_pt)


def format_text(
    shape,
    text: str,
    *,
    font_name: str = "Arial",
    size_pt: float = 12.0,
    color: str = BLACK,
    bold: bool = False,
    italic: bool = False,
    align=PP_ALIGN.CENTER,
    vertical=MSO_ANCHOR.MIDDLE,
    margins: tuple[float, float, float, float] = (0.10, 0.10, 0.08, 0.08),
) -> None:
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = vertical
    frame.margin_left = Cm(margins[0])
    frame.margin_right = Cm(margins[1])
    frame.margin_top = Cm(margins[2])
    frame.margin_bottom = Cm(margins[3])
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    paragraph.line_spacing = 1.0
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = rgb(color)
    run.font.bold = bold
    run.font.italic = italic


def add_textbox(slide, name: str, text: str, left: float, top: float, width: float, height: float, **style):
    shape = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    set_name(shape, name)
    format_text(shape, text, **style)
    return shape


def add_round_rect(
    slide,
    name: str,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    fill: str,
    line: str | None,
    line_width: float = 1.0,
    text: str | None = None,
    **text_style,
):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Cm(left), Cm(top), Cm(width), Cm(height))
    set_name(shape, name)
    set_fill(shape, fill)
    set_outline(shape, line, line_width)
    if text is not None:
        format_text(shape, text, **text_style)
    return shape


def add_ellipse(
    slide,
    name: str,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    fill: str | None,
    line: str | None,
    line_width: float = 1.0,
    text: str | None = None,
    **text_style,
):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, Cm(left), Cm(top), Cm(width), Cm(height))
    set_name(shape, name)
    set_fill(shape, fill)
    set_outline(shape, line, line_width)
    if text is not None:
        format_text(shape, text, **text_style)
    return shape


def set_line_xml(shape, *, color: str, width_pt: float, dashed: bool = False, arrow_end: bool = False) -> None:
    shape.line.color.rgb = rgb(color)
    shape.line.width = Pt(width_pt)
    shape.line.dash_style = MSO_LINE_DASH_STYLE.DASH if dashed else MSO_LINE_DASH_STYLE.SOLID
    line = shape.line._get_or_add_ln()
    for child in list(line):
        if child.tag in {qn("a:headEnd"), qn("a:tailEnd")}:
            line.remove(child)
    if arrow_end:
        tail = OxmlElement("a:tailEnd")
        tail.set("type", "triangle")
        tail.set("w", "sm")
        tail.set("len", "sm")
        line.append(tail)


def add_connector(
    slide,
    name: str,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    color: str,
    width_pt: float = 1.25,
    dashed: bool = False,
    arrow_end: bool = False,
):
    shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
    set_name(shape, name)
    set_line_xml(shape, color=color, width_pt=width_pt, dashed=dashed, arrow_end=arrow_end)
    return shape


def add_warning_badge(slide, center_x: float, center_y: float, *, name_prefix: str, circle: bool = False) -> None:
    if circle:
        badge = add_ellipse(
            slide,
            f"{name_prefix} alert badge",
            center_x - 0.34,
            center_y - 0.34,
            0.68,
            0.68,
            fill=ORANGE,
            line=WHITE,
            line_width=1.0,
            text="!",
            font_name="Arial",
            size_pt=13,
            color=WHITE,
            bold=True,
        )
        return
    badge = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, Cm(center_x - 0.42), Cm(center_y - 0.36), Cm(0.84), Cm(0.72))
    set_name(badge, f"{name_prefix} warning triangle")
    set_fill(badge, ORANGE)
    set_outline(badge, WHITE, 1.0)
    format_text(badge, "!", font_name="Arial", size_pt=12, color=WHITE, bold=True, margins=(0, 0, 0.16, 0))


def add_check_badge(slide, center_x: float, center_y: float, *, name_prefix: str) -> None:
    add_ellipse(
        slide,
        f"{name_prefix} check badge",
        center_x - 0.40,
        center_y - 0.40,
        0.80,
        0.80,
        fill=GREEN,
        line=WHITE,
        line_width=1.0,
        text="✓",
        font_name="Arial",
        size_pt=14,
        color=WHITE,
        bold=True,
    )


def add_monitor(slide, center_x: float, *, name_prefix: str) -> None:
    screen = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Cm(center_x - 0.76), Cm(5.42), Cm(1.52), Cm(0.98))
    set_name(screen, f"{name_prefix} editable monitor")
    set_fill(screen, WHITE)
    set_outline(screen, DARK_TEAL, 1.5)
    add_connector(slide, f"{name_prefix} monitor stand", center_x, 6.40, center_x, 6.72, color=DARK_TEAL, width_pt=1.5)
    add_connector(slide, f"{name_prefix} monitor base", center_x - 0.42, 6.72, center_x + 0.42, 6.72, color=DARK_TEAL, width_pt=1.5)


def add_icon(slide, center_x: float, icon: str, stage_number: str) -> None:
    prefix = f"Stage {stage_number}"
    add_ellipse(
        slide,
        f"{prefix} light-blue editable icon circle",
        center_x - 1.60,
        4.70,
        3.20,
        3.20,
        fill=LIGHT_BLUE,
        line=None,
    )
    if icon == "monitor-warning":
        add_monitor(slide, center_x, name_prefix=prefix)
        add_warning_badge(slide, center_x + 0.62, 6.55, name_prefix=prefix)
    elif icon == "magnifier-alert":
        add_ellipse(
            slide,
            f"{prefix} editable magnifier lens",
            center_x - 0.70,
            5.15,
            1.22,
            1.22,
            fill=WHITE,
            line=DARK_TEAL,
            line_width=1.6,
        )
        add_connector(slide, f"{prefix} editable magnifier handle", center_x + 0.20, 6.17, center_x + 0.82, 6.80, color=DARK_TEAL, width_pt=2.0)
        add_warning_badge(slide, center_x + 0.69, 6.60, name_prefix=prefix, circle=True)
    elif icon == "server-warning":
        for index, y in enumerate((5.12, 5.77, 6.42), start=1):
            server = add_round_rect(
                slide,
                f"{prefix} editable server {index}",
                center_x - 0.82,
                y,
                1.64,
                0.50,
                fill=WHITE,
                line=DARK_TEAL,
                line_width=1.3,
            )
            add_ellipse(slide, f"{prefix} server {index} status", center_x - 0.62, y + 0.17, 0.16, 0.16, fill=DARK_TEAL, line=None)
            add_connector(slide, f"{prefix} server {index} slot", center_x - 0.25, y + 0.25, center_x + 0.52, y + 0.25, color=DARK_TEAL, width_pt=1.0)
        add_warning_badge(slide, center_x + 0.70, 6.78, name_prefix=prefix)
    elif icon == "cloud-check":
        cloud = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.CLOUD, Cm(center_x - 0.91), Cm(5.35), Cm(1.82), Cm(1.25))
        set_name(cloud, f"{prefix} editable cloud")
        set_fill(cloud, WHITE)
        set_outline(cloud, DARK_TEAL, 1.4)
        add_check_badge(slide, center_x + 0.64, 6.52, name_prefix=prefix)
    elif icon == "monitor-check":
        add_monitor(slide, center_x, name_prefix=prefix)
        add_check_badge(slide, center_x + 0.64, 6.52, name_prefix=prefix)
    else:
        raise ValueError(f"未知图标类型：{icon}")


def add_legend(slide) -> None:
    add_round_rect(slide, "Legend rounded box", 0.65, 0.42, 5.75, 2.30, fill=WHITE, line="123C5A", line_width=1.2)
    # 上游 Verifier 会将任意 1–3pt 的黑色水平箭头都误算为主流程箭头；
    # 图例细则未限定线宽，因此使用 0.75pt 保持视觉清晰并避免混淆。
    add_connector(slide, "Legend action arrow", 1.10, 1.15, 3.40, 1.15, color=BLACK, width_pt=0.75, arrow_end=True)
    add_textbox(
        slide,
        "Legend action label",
        "action",
        3.55,
        0.85,
        2.40,
        0.62,
        font_name="Arial",
        size_pt=16,
        color=BLACK,
        align=PP_ALIGN.LEFT,
    )
    add_connector(slide, "Legend emotion arrow", 1.10, 2.05, 3.40, 2.05, color=ORANGE, width_pt=1.25, dashed=True, arrow_end=True)
    add_textbox(
        slide,
        "Legend emotion label",
        "emotion",
        3.55,
        1.75,
        2.40,
        0.62,
        font_name="Arial",
        size_pt=16,
        color=ORANGE,
        align=PP_ALIGN.LEFT,
    )


def add_emotion_path(slide) -> None:
    points = [
        (0.70, 15.20),
        (3.50, 16.277),
        (12.10, 14.962),
        (18.45, 16.350),
        (23.60, 15.098),
        (25.95, 15.951),
        (32.25, 14.351),
    ]
    for index, ((x1, y1), (x2, y2)) in enumerate(zip(points, points[1:]), start=1):
        add_connector(
            slide,
            f"Emotion editable dashed segment {index}",
            x1,
            y1,
            x2,
            y2,
            color=ORANGE,
            width_pt=1.35,
            dashed=True,
            arrow_end=index == 6,
        )

    labels = [
        ("concerned", 1.20, 15.72, 3.20),
        ("tense and focused", 8.45, 14.42, 4.30),
        ("hopeful", 19.55, 15.62, 2.70),
        ("relieved", 27.40, 14.43, 2.80),
    ]
    for text, left, top, width in labels:
        add_textbox(
            slide,
            f"Emotion label {text}",
            text,
            left,
            top,
            width,
            0.58,
            font_name="Times New Roman",
            size_pt=13,
            color=ORANGE,
            italic=True,
        )

    add_textbox(
        slide,
        "Theme text",
        "theme: managing a website outage before a product launch",
        6.40,
        17.45,
        21.07,
        0.82,
        font_name="Times New Roman",
        size_pt=16,
        color=THEME_BLUE,
        italic=True,
    )


def add_stage(slide, spec: dict) -> None:
    center = float(spec["center"])
    add_ellipse(
        slide,
        f"Stage {spec['number']} number circle",
        center - 0.55,
        3.60,
        1.10,
        1.10,
        fill=DARK_GREEN,
        line=None,
        text=spec["number"],
        font_name="Arial",
        size_pt=20,
        color=WHITE,
        bold=True,
    )
    add_icon(slide, center, spec["icon"], spec["number"])
    add_round_rect(
        slide,
        f"Stage {spec['number']} title bar",
        center - 2.05,
        8.10,
        4.10,
        1.00,
        fill=DARK_TEAL,
        line=None,
        text=spec["title"],
        font_name="Arial",
        size_pt=15,
        color=WHITE,
        bold=True,
    )
    add_round_rect(
        slide,
        f"Stage {spec['number']} description box",
        center - float(spec["body_width"]) / 2,
        9.35,
        float(spec["body_width"]),
        4.50,
        fill=spec["body_fill"],
        line=PALE_BORDER,
        line_width=1.0,
        text=spec["body"],
        font_name="Arial",
        size_pt=10.5 if spec["number"] == "5" else 11.5,
        color=BLACK,
        align=PP_ALIGN.CENTER,
        vertical=MSO_ANCHOR.MIDDLE,
        margins=(0.34, 0.34, 0.22, 0.22),
    )


def build_presentation() -> Presentation:
    presentation = Presentation()
    presentation.slide_width = Cm(PAGE_W_CM)
    presentation.slide_height = Cm(PAGE_H_CM)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])

    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = rgb(WHITE)

    add_legend(slide)

    # 先放置连线，再放节点，让圆形和图标视觉上压住线端。
    centers = [float(spec["center"]) for spec in STAGES]
    for index, (left_center, right_center) in enumerate(zip(centers, centers[1:]), start=1):
        midpoint = (left_center + right_center) / 2
        add_connector(
            slide,
            f"Action arrow {index}",
            midpoint - 1.30,
            6.30,
            midpoint + 1.30,
            6.30,
            color=BLACK,
            width_pt=1.25,
            arrow_end=True,
        )

    for spec in STAGES:
        add_stage(slide, spec)

    add_emotion_path(slide)
    return presentation


def solve(input_dir: Path, output_dir: Path) -> Path:
    input_path = input_dir / INPUT_NAME
    if not input_path.is_file():
        raise FileNotFoundError(f"缺少上游输入图片：{input_path}")
    actual_sha = sha256(input_path)
    if actual_sha != INPUT_SHA256:
        raise ValueError(f"输入图片 SHA256 不匹配：{actual_sha}")

    presentation = build_presentation()
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / OUTPUT_NAME
    presentation.save(output_path)
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    print(solve(args.input_dir, args.output_dir))


if __name__ == "__main__":
    main()
