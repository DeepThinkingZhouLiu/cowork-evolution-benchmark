#!/usr/bin/env python3
"""逐页人工转录并改造 OmegaUse-OfficeVal officeval_006。

这不是通用题目生成器。脚本只重放这一道真实上游任务：读取原 PDF 和 Word
模板，把人工核对过的 21 道题及参考答案写成可编辑 DOCX；数学分数、根式使用
原生 OMML，对应的两幅几何图从原 PDF 页面裁取。
"""

from __future__ import annotations

import argparse
import copy
import subprocess
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from PIL import Image


PDF_NAME = "2026年3月17日高中数学作业.pdf"
TEMPLATE_NAME = "高一年级第一次数学周练111 (2).docx"
OUTPUT_NAME = "2026年3月17日高中数学作业_可编辑版.docx"


def set_run_font(run, *, east_asia="宋体", latin="Times New Roman", size=10.5, bold=False):
    run.font.name = latin
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.get_or_add_rFonts()
    fonts.set(qn("w:eastAsia"), east_asia)
    fonts.set(qn("w:ascii"), latin)
    fonts.set(qn("w:hAnsi"), latin)
    sz = rpr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rpr.append(sz)
    sz.set(qn("w:val"), str(int(round(size * 2))))
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(int(round(size * 2))))
    return run


def format_paragraph(paragraph, *, line=1.34, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     before=0, after=0, left=0, first=0):
    fmt = paragraph.paragraph_format
    fmt.alignment = align
    fmt.line_spacing = line
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.left_indent = Cm(left)
    fmt.first_line_indent = Cm(first)
    return paragraph


def body_paragraph(document, text="", *, before=0, after=0, left=0, first=0):
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, before=before, after=after, left=left, first=first)
    if text:
        set_run_font(paragraph.add_run(text))
    return paragraph


def body_run(paragraph, text, *, bold=False):
    return set_run_font(paragraph.add_run(text), bold=bold)


def section_heading(document, text):
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, line=1.35, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     before=2, after=1)
    set_run_font(paragraph.add_run(text), bold=True)
    return paragraph


def math_run(text):
    run = OxmlElement("m:r")
    rpr = OxmlElement("m:rPr")
    sty = OxmlElement("m:sty")
    sty.set(qn("m:val"), "p")
    rpr.append(sty)
    run.append(rpr)
    wrpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Cambria Math")
    fonts.set(qn("w:hAnsi"), "Cambria Math")
    fonts.set(qn("w:eastAsia"), "宋体")
    wrpr.append(fonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "21")
    wrpr.append(sz)
    run.append(wrpr)
    value = OxmlElement("m:t")
    value.text = str(text)
    run.append(value)
    return run


def _parts(value):
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


def _append(container, value):
    for item in _parts(value):
        container.append(math_run(item) if isinstance(item, (str, int, float)) else item)


def frac(numerator, denominator):
    node = OxmlElement("m:f")
    props = OxmlElement("m:fPr")
    kind = OxmlElement("m:type")
    kind.set(qn("m:val"), "bar")
    props.append(kind)
    node.append(props)
    num = OxmlElement("m:num")
    den = OxmlElement("m:den")
    _append(num, numerator)
    _append(den, denominator)
    node.extend([num, den])
    return node


def radical(value):
    node = OxmlElement("m:rad")
    props = OxmlElement("m:radPr")
    hide = OxmlElement("m:degHide")
    hide.set(qn("m:val"), "1")
    props.append(hide)
    node.append(props)
    degree = OxmlElement("m:deg")
    degree.append(math_run(""))
    expression = OxmlElement("m:e")
    _append(expression, value)
    node.extend([degree, expression])
    return node


def superscript(base, exponent):
    node = OxmlElement("m:sSup")
    expr = OxmlElement("m:e")
    sup = OxmlElement("m:sup")
    _append(expr, base)
    _append(sup, exponent)
    node.extend([expr, sup])
    return node


def subscript(base, sub):
    node = OxmlElement("m:sSub")
    expr = OxmlElement("m:e")
    sub_el = OxmlElement("m:sub")
    _append(expr, base)
    _append(sub_el, sub)
    node.extend([expr, sub_el])
    return node


def equation(paragraph, *values):
    node = OxmlElement("m:oMath")
    for value in values:
        _append(node, value)
    paragraph._p.append(node)
    return node


def add_mixed(paragraph, *segments):
    """字符串是普通可编辑正文；其他对象作为原生 OMML 插入。"""
    for segment in segments:
        if isinstance(segment, str):
            body_run(paragraph, segment)
        else:
            equation(paragraph, segment)
    return paragraph


def set_cell_margins(cell, top=20, start=55, bottom=20, end=55):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge, width in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        element = tc_mar.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            tc_mar.append(element)
        element.set(qn("w:w"), str(width))
        element.set(qn("w:type"), "dxa")


def remove_table_borders(table):
    props = table._tbl.tblPr
    borders = props.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        props.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = borders.find(qn(f"w:{edge}"))
        if element is None:
            element = OxmlElement(f"w:{edge}")
            borders.append(element)
        element.set(qn("w:val"), "nil")


def cell_mixed(cell, *segments, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph = cell.paragraphs[0]
    paragraph.clear()
    format_paragraph(paragraph, line=1.0, align=align)
    for segment in segments:
        if isinstance(segment, str):
            set_run_font(paragraph.add_run(segment), bold=bold)
        else:
            equation(paragraph, copy.deepcopy(segment))
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(cell)
    return paragraph


def options_table(document, options, *, columns=4):
    rows = (len(options) + columns - 1) // columns
    table = document.add_table(rows=rows, cols=columns)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    remove_table_borders(table)
    width = 19.8 / columns
    for index, option in enumerate(options):
        row, col = divmod(index, columns)
        cell = table.cell(row, col)
        cell.width = Cm(width)
        if isinstance(option, str):
            cell_mixed(cell, option)
        else:
            cell_mixed(cell, *option)
    return table


def page_break_sentinel(document):
    paragraph = body_paragraph(document, "\u2060")
    paragraph.runs[0].add_break(WD_BREAK.PAGE)
    return paragraph


def add_break(paragraph):
    run = body_run(paragraph, "\u2060")
    run.add_break(WD_BREAK.PAGE)


def add_picture_paragraph(document, image_path, width_cm):
    paragraph = document.add_paragraph()
    format_paragraph(paragraph, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    paragraph.add_run().add_picture(str(image_path), width=Cm(width_cm))
    return paragraph


def crop_source_figures(pdf_path: Path, temp_root: Path):
    prefix2 = temp_root / "page2"
    prefix3 = temp_root / "page3"
    for page, prefix in ((2, prefix2), (3, prefix3)):
        subprocess.run(
            ["pdftoppm", "-png", "-r", "150", "-f", str(page), "-l", str(page),
             "-singlefile", str(pdf_path), str(prefix)],
            check=True, capture_output=True, timeout=120,
        )
    parabola = temp_root / "parabola.png"
    triangle = temp_root / "triangle.png"
    Image.open(prefix2.with_suffix(".png")).crop((480, 210, 790, 490)).save(parabola)
    Image.open(prefix3.with_suffix(".png")).crop((460, 330, 820, 625)).save(triangle)
    return parabola, triangle


def clear_body(document):
    body = document._element.body
    section_properties = body.sectPr
    for child in list(body):
        if child is not section_properties:
            body.remove(child)


def build_exam(document, parabola_path, triangle_path):
    title = document.add_paragraph()
    format_paragraph(title, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_run_font(title.add_run("2026年3月17日高中数学作业"), east_asia="黑体",
                 size=18, bold=True)

    school = document.add_paragraph()
    format_paragraph(school, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER, after=1)
    set_run_font(school.add_run(
        "学校：____________　姓名：____________　班级：____________　考号：____________"
    ), size=10.5, bold=True)

    section_heading(document, "一、单选题")

    p = body_paragraph(document)
    add_mixed(p, "1．已知集合 A={x | log₂(x−1)<2}，B={x | x²−5x+6≤0}，则 A∩B=（　）")
    options_table(document, ["A．(1,5)", "B．(2,3)", "C．[2,3]", "D．(1,2]∪[3,5)"])

    p = body_paragraph(document)
    add_mixed(p, "2．函数 f(x)=aˣ(a>0，a≠1) 的图象经过点 (2,9)，则 f(−1)+logₐ27=（　）")
    options_table(document, [
        ("A．", frac("3", "4")), ("B．", frac("10", "3")), "C．4", ("D．", frac("28", "3")),
    ])

    p = body_paragraph(document)
    add_mixed(p, "3．已知 sin α=", frac("3", "5"), "，且 α∈(", frac("π", "2"), "，π)，则 cos(α+",
              frac("π", "3"), ")=（　）")
    sqrt3 = radical("3")
    options_table(document, [
        ("A．−", frac(["4+3", copy.deepcopy(sqrt3)], "10")),
        ("B．", frac(["4−3", copy.deepcopy(sqrt3)], "10")),
        ("C．", frac(["3", copy.deepcopy(sqrt3), "−4"], "10")),
        ("D．", frac(["4+3", copy.deepcopy(sqrt3)], "10")),
    ])

    p = body_paragraph(document, "4．等差数列 {aₙ} 中，a₁=3，a₈=24，则 S₁₀=（　）")
    options_table(document, ["A．135", "B．150", "C．160", "D．165"])

    p = body_paragraph(document, "5．圆 C：x²+y²−4x+2y−4=0 的圆心到直线 3x−4y+1=0 的距离为（　）")
    options_table(document, [("A．", frac("7", "5")), "B．2", ("C．", frac("11", "5")), "D．3"])

    p = body_paragraph(document, "6．袋中有 3 个红球、2 个蓝球、1 个白球，除颜色外完全相同。从中任取 2 个球，恰好取到 1 个红球和 1 个蓝球的概率为（　）")
    options_table(document, [
        ("A．", frac("1", "5")), ("B．", frac("2", "5")),
        ("C．", frac("1", "2")), ("D．", frac("3", "5")),
    ])

    p = body_paragraph(document, "7．若不等式 x²−2mx+4>0 对一切 x∈R 恒成立，则 m 的取值范围是（　）")
    options_table(document, ["A．(−2,2)", "B．[−2,2]", "C．(−∞,−2)∪(2,+∞)", "D．R"], columns=2)

    p = body_paragraph(document, "8．在正方体 ABCD−A₁B₁C₁D₁ 中，若 AB=1，则异面直线 AB₁ 与 CD₁ 所成的角为（　）")
    options_table(document, ["A．30°", "B．45°", "C．90°", "D．120°"])

    p = body_paragraph(document, "9．函数 f(x)=x³−3x 在区间 [−2,2] 上的最大值为（　）")
    options_table(document, ["A．0", "B．2", "C．4", "D．6"])
    page_break_sentinel(document)

    p = body_paragraph(document, "10．如图，抛物线 y=x²−2x−3 与 x 轴交于 A、B 两点，点 B 在点 A 的右侧，P 为抛物线顶点，则 △ABP 的面积为（　）")
    add_picture_paragraph(document, parabola_path, 3.7)
    options_table(document, ["A．4", "B．6", "C．7", "D．8"])

    section_heading(document, "二、多选题")
    p = body_paragraph(document, "11．下列函数在 (0,+∞) 上单调递增的是（　）")
    options_table(document, ["A．y=ln x", "B．y=x+1/x", "C．y=x²", "D．y=2ˣ"])

    p = body_paragraph(document, "12．已知等比数列 {aₙ} 满足 a₁=2，a₃=18，且公比 q>0。下列结论正确的是（　）")
    options_table(document, ["A．q=3", "B．a₅=162", "C．S₄=80", "D．a₂+a₄=48"])

    p = body_paragraph(document)
    add_mixed(p, "13．直线 y=kx+1 与圆 x²+y²=", frac("1", "4"), " 有两个不同交点。下列 k 的取值可行的是（　）")
    options_table(document, ["A．0", "B．1", "C．2", "D．−3"])

    p = body_paragraph(document, "14．同时掷两枚质地均匀的骰子，设事件 A 为“点数和为 7”，事件 B 为“至少一枚骰子点数为 6”，事件 C 为“两枚骰子的点数相同”。下列说法正确的是（　）")
    options_table(document, [
        ("A．P(A)=", frac("1", "6")), ("B．P(B)=", frac("11", "36")),
        "C．A 与 C 互斥", "D．A 与 B 相互独立",
    ], columns=2)

    section_heading(document, "三、填空题")
    p = body_paragraph(document)
    add_mixed(p, "15．函数 f(x)=", radical("x−1"), "+", frac("1", "3−x"), " 的定义域为________________。")
    p = body_paragraph(document, "16．已知 tan α=2，则 sin 2α=________________。")
    p = body_paragraph(document)
    add_mixed(p, "17．若双曲线 ", frac("x²", "a²"), "−", frac("y²", "9"), "=1(a>0) 的一条渐近线方程为 y=", frac("3", "2"), "x，则 a=________________。")

    section_heading(document, "四、解答题")
    p = body_paragraph(document, "18．已知数列 {aₙ} 满足 a₁=1，aₙ₊₁=2aₙ+1(n∈N*)。")
    add_break(p)

    body_paragraph(document, "（1）证明数列 {aₙ+1} 是等比数列；", left=0.9)
    body_paragraph(document, "（2）求 aₙ 及前 n 项和 Sₙ。", left=0.9)

    p = body_paragraph(document, "19．如图，在等腰三角形 ABC 中，AB=AC=5，BC=6，D 为 BC 的中点，点 E 在 AD 上且 DE=1。过 E 作 MN∥BC，分别交 AB、AC 于 M、N。")
    add_picture_paragraph(document, triangle_path, 4.2)
    body_paragraph(document, "（1）求 AD 的长；", left=0.9)
    body_paragraph(document, "（2）求 MN 的长；", left=0.9)
    body_paragraph(document, "（3）求 △AMN 的面积。", left=0.9)

    p = body_paragraph(document)
    add_mixed(p, "20．已知椭圆 C：", frac("x²", "4"), "+y²=1，直线 l：y=x+m。")
    body_paragraph(document, "（1）求直线 l 与椭圆 C 相切时 m 的值；", left=0.9)
    body_paragraph(document, "（2）当 m=1 时，求直线 l 被椭圆 C 截得的弦长。", left=0.9)

    body_paragraph(document, "21．已知函数 f(x)=x³−3ax+2 在 x=1 处取得极小值。")
    body_paragraph(document, "（1）求实数 a 的值；", left=0.9)
    p = body_paragraph(document, "（2）求 f(x) 在区间 [−2,3] 上的最大值与最小值。", left=0.9)
    add_break(p)


def answer_table(document, question_numbers, answers, widths=None):
    columns = len(question_numbers) + 1
    table = document.add_table(rows=2, cols=columns)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    remove_table_borders(table)
    values = [["题号", *question_numbers], ["答案", *answers]]
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            value = values[row_index][col_index]
            if isinstance(value, (list, tuple)):
                cell_mixed(cell, *value, align=WD_ALIGN_PARAGRAPH.CENTER)
            else:
                cell_mixed(cell, str(value), align=WD_ALIGN_PARAGRAPH.CENTER)
            cell.width = Cm(widths[col_index] if widths else (18.5 / columns))
    return table


def answer_title(document):
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(table)
    cell_mixed(table.cell(0, 0), "《2026年3月17日高中数学作业》参考答案",
               bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    return table


def build_answers(document):
    answer_title(document)
    answer_table(document, [str(i) for i in range(1, 11)], list("CBADC BACBD".replace(" ", "")))
    answer_table(document, ["11", "12", "13", "14", "15", "16", "17"],
                 ["ACD", "ABC", "CD", "ABC", "[1,3)∪(3,+∞)", (frac("4", "5"),), "2"],
                 widths=[1.5, 1.7, 1.7, 1.7, 1.7, 4.4, 1.7, 1.7])

    p = body_paragraph(document, "1．A 中 1<x<5，B 中 2≤x≤3，所以 A∩B=[2,3]。故选 C。")

    p = body_paragraph(document)
    add_mixed(p, "2．由 a²=9 且 a>0 得 a=3。于是 f(−1)=3⁻¹=", frac("1", "3"), "，log₃27=3，所求为 ", frac("10", "3"), "。故选 B。")

    p = body_paragraph(document)
    add_mixed(p, "3．α 在第二象限，故 cos α=−", frac("4", "5"), "。于是 cos(α+", frac("π", "3"), ")=cos α cos ", frac("π", "3"), "−sin α sin ", frac("π", "3"), "=−", frac("2", "5"), "−", frac(["3", radical("3")], "10"), "=−", frac(["4+3", radical("3")], "10"), "。故选 A。")

    p = body_paragraph(document)
    add_mixed(p, "4．由 a₈=a₁+7d=24 得 d=3，所以 a₁₀=30，S₁₀=", frac("10(3+30)", "2"), "=165。故选 D。")

    p = body_paragraph(document)
    add_mixed(p, "5．圆心为 (2,−1)，代入点到直线的距离公式得 d=", frac("11", "5"), "。故选 C。")

    p = body_paragraph(document)
    add_mixed(p, "6．样本总数为 C₆²=15，符合条件的取法有 3×2=6 种，概率为 ", frac("6", "15"), "=", frac("2", "5"), "。故选 B。")

    p = body_paragraph(document, "7．二次函数开口向上，且恒大于 0，需判别式 Δ=(−2m)²−16<0，即 −2<m<2。故选 A。")
    p = body_paragraph(document, "8．取坐标 A(0,0,0)、B₁(1,0,1)、C(1,1,0)、D₁(0,1,1)，则向量 AB₁=(1,0,1)，向量 CD₁=(−1,0,1)，数量积为 0，故夹角为 90°。故选 C。")
    p = body_paragraph(document, "9．f′(x)=3x²−3，临界点为 x=±1。比较 f(−2)=−2，f(−1)=2，f(1)=−2，f(2)=2，最大值为 2。故选 B。")

    p = body_paragraph(document)
    add_mixed(p, "10．抛物线与 x 轴交点为 A(−1,0)、B(3,0)，顶点为 P(1,−4)。所以 AB=4，点 P 到 x 轴的距离为 4，面积为 ", frac("1", "2"), "×4×4=8。故选 D。")

    p = body_paragraph(document, "11．ln x、x²、2ˣ 均在 (0,+∞) 上单调递增；x+1/x 在该区间内不是单调递增。故选 ACD。")
    p = body_paragraph(document, "12．由 a₃=a₁q² 得 18=2q²，又 q>0，故 q=3。于是 a₅=2·3⁴=162，S₄=2(1+3+9+27)=80，a₂+a₄=6+54=60。故选 ABC。")

    p = body_paragraph(document)
    add_mixed(p, "13．圆心到直线 kx−y+1=0 的距离为 ", frac("1", radical("k²+1")), "。有两个不同交点需 ", frac("1", radical("k²+1")), "<", frac("1", "2"), "，即 |k|>", radical("3"), "。可行的是 2 和 −3。故选 CD。")

    p = body_paragraph(document)
    add_mixed(p, "14．点数和为 7 共有 6 种，故 P(A)=", frac("1", "6"), "；至少一枚为 6 共有 11 种，故 P(B)=", frac("11", "36"), "；事件 A 与 C 不可能同时发生，故互斥；P(A∩B)=", frac("2", "36"), "=", frac("1", "18"), "≠P(A)P(B)，故 A 与 B 不独立。故选 ABC。")
    add_break(p)

    p = body_paragraph(document)
    add_mixed(p, "15．根式要求 x−1≥0，分母要求 x≠3，故定义域为 [1,3)∪(3,+∞)。")

    p = body_paragraph(document)
    add_mixed(p, "16．sin 2α=", frac("2tan α", "1+tan²α"), "=", frac("4", "5"), "。")

    p = body_paragraph(document)
    add_mixed(p, "17．双曲线的渐近线为 y=±", frac("3", "a"), "x，由 ", frac("3", "a"), "=", frac("3", "2"), " 得 a=2。")

    p = body_paragraph(document, "18．（1）因为 aₙ₊₁+1=2aₙ+2=2(aₙ+1)，且 a₁+1=2，所以 {aₙ+1} 是首项为 2、公比为 2 的等比数列。")
    p = body_paragraph(document, "（2）由上知 aₙ+1=2ⁿ，所以 aₙ=2ⁿ−1。因此 Sₙ=∑(2ᵏ−1)=2ⁿ⁺¹−n−2。", left=0.9)

    p = body_paragraph(document)
    add_mixed(p, "19．（1）因为 AB=AC，D 为 BC 中点，所以 AD⊥BC，BD=3。由勾股定理得 AD=", radical("5²−3²"), "=4。")
    p = body_paragraph(document)
    add_mixed(p, "（2）DE=1，故 AE=AD−DE=3。又 MN∥BC，所以 △AMN∽△ABC，相似比为 ", frac("AE", "AD"), "=", frac("3", "4"), "，故 MN=", frac("3", "4"), "×6=", frac("9", "2"), "。")
    p = body_paragraph(document)
    add_mixed(p, "（3）S△ABC=", frac("1", "2"), "×6×4=12，面积比为相似比的平方，故 S△AMN=12×(", frac("3", "4"), ")²=", frac("27", "4"), "。")

    p = body_paragraph(document)
    add_mixed(p, "20．（1）联立两式得 5x²+8mx+4m²−4=0。相切时判别式为 0，即 64m²−20(4m²−4)=0，故 m=±", radical("5"), "。")
    p = body_paragraph(document)
    add_mixed(p, "（2）当 m=1 时，方程为 5x²+8x=0，交点横坐标为 0 与 −", frac("8", "5"), "，对应纵坐标为 1 与 −", frac("3", "5"), "。弦长为 ", radical(["(", frac("8", "5"), ")²+(", frac("8", "5"), ")²"]), "=", frac(["8", radical("2")], "5"), "。")

    p = body_paragraph(document, "21．（1）f′(x)=3x²−3a。在 x=1 处取得极小值，故 f′(1)=0，得 a=1。此时 f″(1)=6>0，条件满足。")
    p = body_paragraph(document, "（2）f(x)=x³−3x+2，f′(x)=3x²−3，临界点为 x=−1、1。比较 f(−2)=0，f(−1)=4，f(1)=0，f(3)=20。因此最大值为 20，最小值为 0。")


def normalize_footer(document):
    """保留模板页脚并把页码段落显式设为居中。"""
    for section in document.sections:
        for paragraph in section.footer.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def create_document(input_dir: Path, output_dir: Path):
    pdf_path = input_dir / PDF_NAME
    template_path = input_dir / TEMPLATE_NAME
    if not pdf_path.is_file() or not template_path.is_file():
        raise FileNotFoundError(f"缺少原始输入：{pdf_path} 或 {template_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="officeval-006-") as temp_dir:
        temp_root = Path(temp_dir)
        parabola, triangle = crop_source_figures(pdf_path, temp_root)
        document = Document(template_path)
        clear_body(document)
        build_exam(document, parabola, triangle)
        build_answers(document)
        normalize_footer(document)
        output_path = output_dir / OUTPUT_NAME
        document.save(output_path)
    return output_path


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    print(create_document(args.input_dir, args.output_dir))
