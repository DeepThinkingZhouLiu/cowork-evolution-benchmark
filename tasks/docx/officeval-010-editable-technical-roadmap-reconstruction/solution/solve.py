#!/usr/bin/env python3
"""officeval_010 的逐题参考解法：按来源图逐对象重建删行后的技术路线图。"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm


INPUT_NAME = "技术路线图.docx"
OUTPUT_NAME = "技术路线图_删除第四行_可编辑版.docx"
INPUT_SHA256 = "e48285e337a48ca41ad7918b786a373360df043967f65dcfe58a38ed97f61817"
GREEN = "#17913A"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def text_xml(text: str, *, size_pt: float = 9, bold: bool = False, vertical: bool = False) -> str:
    parts = list(text) if vertical else [text]
    paragraphs = []
    for part in parts:
        paragraphs.append(
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr>'
            '<w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="黑体"/>'
            f'<w:sz w:val="{int(size_pt * 2)}"/><w:szCs w:val="{int(size_pt * 2)}"/>'
            + ('<w:b/>' if bold else '')
            + f'</w:rPr><w:t>{escape(part)}</w:t></w:r></w:p>'
        )
    return '<v:textbox inset="1pt,1pt,1pt,1pt"><w:txbxContent>' + ''.join(paragraphs) + '</w:txbxContent></v:textbox>'


def rect(identifier: str, text: str, x: float, y: float, w: float, h: float, *, fill: str = "#FFFFFF", stroke: str = "#000000", size: float = 9, bold: bool = False, vertical: bool = False, rounded: bool = False, visible: bool = True) -> str:
    tag = "v:roundrect" if rounded else "v:rect"
    style = (
        f"position:absolute;mso-position-horizontal:absolute;mso-position-vertical:absolute;"
        f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
        f"left:{x}cm;top:{y}cm;width:{w}cm;height:{h}cm;z-index:2;v-text-anchor:middle"
    )
    attrs = f'id="{identifier}" style="{style}" fillcolor="{fill}" strokecolor="{stroke}" strokeweight="1pt"'
    if not visible:
        attrs += ' filled="f" stroked="f"'
    return f'<{tag} {attrs}>{text_xml(text, size_pt=size, bold=bold, vertical=vertical)}</{tag}>'


def diamond(identifier: str, text: str, x: float, y: float, w: float, h: float) -> str:
    style = (
        f"position:absolute;mso-position-horizontal:absolute;mso-position-vertical:absolute;"
        f"mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
        f"left:{x}cm;top:{y}cm;width:{w}cm;height:{h}cm;z-index:2;v-text-anchor:middle"
    )
    return (
        f'<v:shape id="{identifier}" coordsize="21600,21600" path="m10800,0l21600,10800,10800,21600,0,10800xe" '
        f'style="{style}" fillcolor="#FFFFFF" strokecolor="#000000" strokeweight="1pt">'
        f'{text_xml(text)}</v:shape>'
    )


def line(identifier: str, x1: float, y1: float, x2: float, y2: float, *, color: str = GREEN, weight: float = 1.0, arrow: bool = False) -> str:
    stroke = f'<v:stroke dashstyle="solid"' + (' endarrow="block" endarrowlength="short" endarrowwidth="narrow"' if arrow else '') + '/>'
    # LibreOffice 需要明确的页面绝对定位，否则会把 from/to 当成段落内坐标，
    # 导致结构检查正确但连线全部挤在页顶。
    style = (
        "position:absolute;mso-position-horizontal:absolute;mso-position-vertical:absolute;"
        "mso-position-horizontal-relative:page;mso-position-vertical-relative:page;z-index:1"
    )
    return (
        f'<v:line id="{identifier}" style="{style}" from="{x1}cm,{y1}cm" to="{x2}cm,{y2}cm" '
        f'strokecolor="{color}" strokeweight="{weight}pt" o:allowincell="f">{stroke}</v:line>'
    )


def build_vml() -> list[str]:
    objects = []
    objects.append(rect("title", "技术路线图", 6.0, 2.0, 9.0, 0.8, size=24, bold=True, visible=False))
    objects.append(line("title-separator", 1.7, 3.4, 19.4, 3.4, color="#A6A6A6", weight=1.25))

    objects.append(rect("intro", "研究导入：明确目标与范围", 7.7, 4.0, 5.6, 0.8))
    second = [(4.25, "政策信息梳理"), (7.55, "资料归集校核"), (10.85, "参照案例对照"), (14.15, "业务流程梳理")]
    for index, (x, text) in enumerate(second, 1):
        objects.append(rect(f"second-{index}", text, x, 5.5, 2.6, 0.7))
    objects.extend([
        line("intro-stub", 10.5, 4.8, 10.5, 5.15),
        line("intro-bus", 5.55, 5.15, 15.45, 5.15),
    ])
    for index, (x, _) in enumerate(second, 1):
        center = x + 1.3
        objects.append(line(f"intro-child-{index}", center, 5.15, center, 5.5, arrow=True))

    objects.append(rect("goal", "需求要点识别与目标分解", 7.25, 7.0, 6.5, 0.9))
    for index, (x, _) in enumerate(second, 1):
        center = x + 1.3
        objects.append(line(f"second-stub-{index}", center, 6.2, center, 6.65))
    objects.append(line("second-bus", 5.55, 6.65, 15.45, 6.65))
    objects.append(line("second-goal", 10.5, 6.65, 10.5, 7.0, arrow=True))

    # 原图中位于此处的三个检查框已按 Query 删除；需求节点直接通向可行性确认。
    objects.append(diamond("feasibility", "可行性确认", 9.1, 9.9, 2.8, 1.8))
    objects.append(line("goal-diamond", 10.5, 7.9, 10.5, 9.9, arrow=True))
    objects.append(rect("improve", "完善资料与方案", 14.0, 10.4, 2.8, 0.8))
    objects.append(line("diamond-improve", 11.9, 10.8, 14.0, 10.8, arrow=True))
    objects.append(line("loop-right", 16.8, 10.8, 17.8, 10.8))
    objects.append(line("loop-up", 17.8, 10.8, 17.8, 5.35))
    objects.append(line("loop-back", 17.8, 5.35, 16.75, 5.35, arrow=True))

    objects.append(rect("judge", "运行环境综合研判", 7.25, 12.35, 6.5, 0.8))
    objects.append(line("diamond-judge", 10.5, 11.7, 10.5, 12.35, arrow=True))

    objects.append(rect("external", "外部条件研判", 4.8, 14.5, 3.3, 0.8))
    objects.append(rect("internal", "内部能力评阅", 12.8, 14.5, 3.3, 0.8))
    objects.append(line("judge-stub", 10.5, 13.15, 10.5, 13.7))
    objects.append(line("judge-bus", 6.45, 13.7, 14.45, 13.7))
    objects.append(line("judge-external", 6.45, 13.7, 6.45, 14.5, arrow=True))
    objects.append(line("judge-internal", 14.45, 13.7, 14.45, 14.5, arrow=True))

    left_boxes = [(4.66, "政策环境扫描"), (6.16, "需求场景识别"), (7.66, "协同路径分析")]
    right_boxes = [(11.94, "资源储备评估"), (13.43, "流程执行能力"), (14.92, "技术支撑能力"), (16.41, "运营协调机制")]
    for index, (x, text) in enumerate(left_boxes + right_boxes, 1):
        objects.append(rect(f"vertical-{index}", text, x, 16.5, 0.68, 2.72, vertical=True))
    objects.append(line("external-bus", 5.0, 15.7, 8.0, 15.7))
    for index, (x, _) in enumerate(left_boxes, 1):
        objects.append(line(f"external-child-{index}", x + 0.35, 15.7, x + 0.35, 16.5, arrow=True))
    objects.append(line("internal-bus", 12.28, 15.7, 16.75, 15.7))
    for index, (x, _) in enumerate(right_boxes, 1):
        objects.append(line(f"internal-child-{index}", x + 0.35, 15.7, x + 0.35, 16.5, arrow=True))

    objects.append(rect("efe", "EFE评阅清单", 4.9, 19.65, 2.8, 0.8))
    objects.append(rect("matrix", "适配矩阵", 9.1, 19.65, 2.8, 0.8))
    objects.append(rect("ife", "IFE评阅清单", 13.6, 19.65, 2.8, 0.8))
    objects.append(line("matrix-left", 9.1, 20.05, 8.7, 20.05, arrow=True))
    objects.append(line("matrix-right", 11.9, 20.05, 12.3, 20.05, arrow=True))
    objects.append(line("demand-efe", 6.5, 19.2, 6.3, 20.05, arrow=True))
    objects.append(line("tech-ife", 15.2, 19.2, 15.0, 20.05, arrow=True))

    objects.append(rect("path", "实施路径定位与方案设计", 7.25, 21.3, 6.5, 0.8))
    objects.append(rect("organization", "组织保障与质量闭环", 7.25, 22.8, 6.5, 0.8))
    objects.append(rect("summary", "归纳与展望", 7.25, 24.3, 6.5, 0.8))
    objects.append(line("matrix-path", 10.5, 20.45, 10.5, 21.3, arrow=True))
    objects.append(line("path-org", 10.5, 22.1, 10.5, 22.8, arrow=True))
    objects.append(line("org-summary", 10.5, 23.6, 10.5, 24.3, arrow=True))

    for identifier, text, y in [("stage-intro", "导入", 4.0), ("stage-check", "核验", 7.7), ("stage-judge", "研判", 12.65), ("stage-form", "形成", 21.05)]:
        objects.append(rect(identifier, text, 2.5, y, 0.8, 2.4, fill="#E7E6E6", stroke="#BFBFBF", bold=True, vertical=True, rounded=True))
    objects.append(rect("note", "注：图中机构、业务数据与流程节点均为虚构化编制，未包含真实联系方式或金额。", 4.8, 26.0, 12.5, 0.6, size=7.5, visible=False))

    return objects


def build_document() -> Document:
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(0)
    section.bottom_margin = Cm(0)
    section.left_margin = Cm(0)
    section.right_margin = Cm(0)
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for object_xml in build_vml():
        run = paragraph.add_run()
        run._r.append(parse_xml(
            '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">'
            + object_xml + '</w:pict>'
        ))
    return document


def solve(input_dir: Path, output_dir: Path) -> Path:
    source = input_dir / INPUT_NAME
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("技术路线图输入缺失或 SHA256 不匹配")
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    build_document().save(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    print(solve(args.input_dir, args.output_dir))


if __name__ == "__main__":
    main()
