#!/usr/bin/env python3
"""officeval_003 的逐题参考解法。

脚本只处理这份真实会议记录：先校验输入文件，再逐行清理来源表格，
不生成新会议内容，也不把参考图嵌入交付物。
"""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


INPUT_NAME = "部门协调会记录 (1).docx"
OUTPUT_NAME = "部门协调会记录_会议内容横线表格版.docx"
INPUT_SHA256 = "3c0d0060f17a8b4177c5ed18c6a54502c7a03d3ac6cfdf96b8dadffaaa597344"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def set_cell_margins(cell, value_dxa: int = 36) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for side in ("top", "left", "bottom", "right"):
        node = margins.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            margins.append(node)
        node.set(qn("w:w"), str(value_dxa))
        node.set(qn("w:type"), "dxa")


def set_bottom_rule(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for side in ("left", "right", "insideH", "insideV"):
        node = borders.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            borders.append(node)
        node.set(qn("w:val"), "nil")
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "000000")


def remove_row(table, row) -> None:
    table._tbl.remove(row._tr)


def format_meeting_table(document: Document) -> None:
    if len(document.tables) != 1:
        raise ValueError(f"预期 1 张会议内容表，实际为 {len(document.tables)} 张")
    table = document.tables[0]

    for row in list(table.rows):
        text = "".join(cell.text for cell in row.cells).strip()
        if not text:
            remove_row(table, row)

    if len(table.rows) != 95:
        raise ValueError(f"清理后应为 95 条非空会议记录，实际为 {len(table.rows)} 条")

    for row in table.rows:
        row.height = Cm(0.58)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        if len(row.cells) != 1:
            raise ValueError("会议内容表的每行必须只有一个贯穿整行的单元格")
        cell = row.cells[0]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_margins(cell)
        set_bottom_rule(cell)

        original_text = cell.text.strip()
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.left_indent = Pt(0)
        paragraph.paragraph_format.right_indent = Pt(0)
        paragraph.paragraph_format.first_line_indent = Pt(0)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run = paragraph.add_run(original_text)
        run.font.name = "宋体"
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
        run.font.size = Pt(10)


def solve(input_dir: Path, output_dir: Path) -> Path:
    source = input_dir / INPUT_NAME
    if not source.is_file():
        raise FileNotFoundError(source)
    actual = sha256(source)
    if actual != INPUT_SHA256:
        raise ValueError(f"输入文件 SHA256 不匹配：{actual}")

    document = Document(source)
    format_meeting_table(document)
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    document.save(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    print(solve(args.input_dir, args.output_dir))


if __name__ == "__main__":
    main()
