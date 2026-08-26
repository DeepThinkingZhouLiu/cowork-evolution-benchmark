#!/usr/bin/env python3
"""officeval_029 的逐题参考解法：按已审读参考图逐格搭建表单。"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


INPUT_NAME = "image_001.jpg"
OUTPUT_NAME = "培训记录表_按图制作_可编辑.docx"
INPUT_SHA256 = "65693cbebf7380c963b8445e9d33feb0cb5d08dd590cea1c360d9d77c75af35f"
BASE_WIDTHS = [2.4, 2.4, 1.45, 1.2, 1.45, 2.4, 2.4, 2.4]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(path.read_bytes())
    return digest.hexdigest()


def set_cell_width(cell, width_cm: float) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(round(width_cm * 567)))
    tc_w.set(qn("w:type"), "dxa")
    cell.width = Cm(width_cm)


def set_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "4")
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")
        borders.append(node)


def set_grid(table) -> None:
    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in BASE_WIDTHS:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(round(width * 567)))
        grid.append(column)


def merge_row(row, spans: list[int]):
    cells = row.cells
    cursor = 0
    merged = []
    for span in spans:
        start = cells[cursor]
        end = cells[cursor + span - 1]
        merged.append(start.merge(end) if span > 1 else start)
        cursor += span
    return merged


def format_cell(cell, text: str) -> None:
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.font.name = "仿宋"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "仿宋")
    run.font.size = Pt(15)


def build_document() -> Document:
    document = Document()
    section = document.sections[0]
    section.start_type = WD_SECTION_START.NEW_PAGE
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    table = document.add_table(rows=8, cols=8)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_grid(table)
    set_table_borders(table)

    row1 = list(table.rows[0].cells)
    for cell, width, text in zip(row1, BASE_WIDTHS, ["员工姓名", "", "所属部门", "", "性别", "", "任职岗位", ""]):
        set_cell_width(cell, width)
        format_cell(cell, text)

    row2 = merge_row(table.rows[1], [1, 7])
    for cell, width, text in zip(row2, [2.4, 13.7], ["培训日期", ""]):
        set_cell_width(cell, width)
        format_cell(cell, text)

    row3 = merge_row(table.rows[2], [1, 1, 1, 1, 2, 2])
    for cell, width, text in zip(row3, [2.4, 2.6, 2.85, 2.85, 2.65, 2.65], ["培训目标", "", "培训项目", "", "学习时长", ""]):
        set_cell_width(cell, width)
        format_cell(cell, text)

    row4 = merge_row(table.rows[3], [1, 1, 2, 4])
    for cell, width, text in zip(row4, [2.4, 2.6, 2.85, 8.25], ["考核结果", "", "能力评定", ""]):
        set_cell_width(cell, width)
        format_cell(cell, text)

    labels = ["评定方式", "学习完成\n情况", "直属\n主管\n意见", "人力\n资源\n意见"]
    heights = [1.4, 2.8, 4.2, 4.2]
    for row, label, height in zip(table.rows[4:], labels, heights):
        cells = merge_row(row, [1, 7])
        set_cell_width(cells[0], 2.4)
        set_cell_width(cells[1], 13.7)
        format_cell(cells[0], label)
        format_cell(cells[1], "")
        row.height = Cm(height)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    for row in table.rows[:4]:
        row.height = Cm(1.4)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    return document


def solve(input_dir: Path, output_dir: Path) -> Path:
    source = input_dir / INPUT_NAME
    if not source.is_file() or sha256(source) != INPUT_SHA256:
        raise ValueError("参考图片缺失或 SHA256 不匹配")
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
