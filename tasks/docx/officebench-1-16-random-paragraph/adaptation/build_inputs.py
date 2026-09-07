#!/usr/bin/env python3
"""构造本题专用的文案审校模板和批准改写票据。

所有内容事实都来自 OfficeBench 冻结的 random_paragraph.txt；票据只增加编辑操作和不得扩写的边界。
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


APPROVED = [
    (
        "S1",
        "A small café offered a tranquil pause in the heart of the bustling city.",
        "Subject front-loading",
        "",
        "",
    ),
    (
        "S2",
        "Vintage furniture and soft lighting made its cozy interior inviting to locals and tourists alike.",
        "Clause tightening",
        "locals and tourists alike",
        "Audience evidence comes from source sentence S2; retain both audience groups.",
    ),
    (
        "S3",
        "Freshly brewed coffee and baked pastries scented the room, welcoming everyone who entered.",
        "Sensory consolidation",
        "Freshly brewed coffee and baked pastries",
        "Sensory evidence combines the two aromas named in source sentence S3; no additional menu items are approved.",
    ),
    (
        "S4",
        "At mismatched tables, patrons talked, reflected, and sipped their drinks.",
        "Parallel compression",
        "",
        "",
    ),
    (
        "S5",
        "The city buzzed outside; within the café, time seemed to slow and offer a brief respite.",
        "Contrast compression",
        "time seemed to slow",
        "This figurative claim is retained from source sentence S5 and must not be restated as a measurable service promise.",
    ),
]


def split_sentences(text: str) -> list[str]:
    return [item.strip() for item in re.split(r"(?<=[.!?])\s+", text.strip()) if item.strip()]


def shade(cell, fill: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(node)


def repeat_header(row) -> None:
    node = OxmlElement("w:tblHeader")
    node.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(node)


def keep_row(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def set_cell_text(cell, value: str, bold: bool = False, size: float = 9.0) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(value)
    run.bold = bold
    run.font.size = Pt(size)


def add_page_fields(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.add_run("Controlled copy • Page ")
    for instruction in ("PAGE", "NUMPAGES"):
        if instruction == "NUMPAGES":
            paragraph.add_run(" of ")
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), instruction)
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = "1"
        run.append(text)
        field.append(run)
        paragraph._p.append(field)


def configure_document(document: Document) -> None:
    styles = document.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    styles["Title"].font.name = "Aptos Display"
    styles["Title"].font.color.rgb = RGBColor(31, 78, 121)
    styles["Heading 1"].font.name = "Aptos Display"
    styles["Heading 1"].font.color.rgb = RGBColor(31, 78, 121)

    working = styles.add_style("Working Sentence", 1)
    working.font.name = "Aptos"
    working.font.size = Pt(10.5)
    working.paragraph_format.left_indent = Inches(0.25)
    working.paragraph_format.first_line_indent = Inches(-0.25)
    working.paragraph_format.space_after = Pt(7)

    note = styles.add_style("Evidence Note", 1)
    note.font.name = "Aptos"
    note.font.size = Pt(10)
    note.paragraph_format.space_after = Pt(6)
    note.paragraph_format.line_spacing = 1.08


def build_template(source: str, target: Path) -> None:
    sentences = split_sentences(source)
    if len(sentences) != 5:
        raise ValueError(f"预期 5 句源文，实际 {len(sentences)}")

    document = Document()
    configure_document(document)
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.header.paragraphs[0].text = "NORTHSTAR COPY DESK  |  CAFÉ PROFILE"
    section.header.paragraphs[0].runs[0].font.size = Pt(8)
    add_page_fields(section.footer.paragraphs[0])

    title = document.add_heading("Café Profile Editorial Review", 0)
    title.paragraph_format.space_after = Pt(3)
    subtitle = document.add_paragraph("Controlled source, approved working copy, and sentence-level evidence")
    subtitle.style = document.styles["Subtitle"]

    metadata = document.add_table(rows=3, cols=2)
    metadata.style = "Table Grid"
    values = [
        ("Work item", "CP-016"),
        ("Authority", "random_paragraph.txt"),
        ("Review state", "Awaiting approved-copy application"),
    ]
    for row, (label, value) in zip(metadata.rows, values):
        set_cell_text(row.cells[0], label, bold=True)
        set_cell_text(row.cells[1], value)
        shade(row.cells[0], "D9EAF7")
        keep_row(row)

    document.add_heading("Controlled Source", level=1)
    document.add_paragraph(
        "This block is the fact authority. Preserve it verbatim and use it to check every approved revision.",
        style="Evidence Note",
    )
    source_table = document.add_table(rows=1, cols=1)
    source_table.style = "Table Grid"
    set_cell_text(source_table.cell(0, 0), source, size=10.0)
    shade(source_table.cell(0, 0), "F2F2F2")
    keep_row(source_table.rows[0])

    document.add_heading("Approved Working Copy", level=1)
    document.add_paragraph(
        "Apply the approved ticket as five tracked replacements. Keep each sentence ID outside the tracked range.",
        style="Evidence Note",
    )
    for index, sentence in enumerate(sentences, start=1):
        paragraph = document.add_paragraph(style="Working Sentence")
        label = paragraph.add_run(f"S{index} | ")
        label.bold = True
        paragraph.add_run(sentence)

    landscape = document.add_section(WD_SECTION.NEW_PAGE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = section.page_height, section.page_width
    landscape.top_margin = Inches(0.55)
    landscape.bottom_margin = Inches(0.55)
    landscape.left_margin = Inches(0.55)
    landscape.right_margin = Inches(0.55)
    landscape.header.is_linked_to_previous = False
    landscape.footer.is_linked_to_previous = False
    landscape.header.paragraphs[0].text = "NORTHSTAR COPY DESK  |  SENTENCE EVIDENCE REGISTER"
    landscape.header.paragraphs[0].runs[0].font.size = Pt(8)
    add_page_fields(landscape.footer.paragraphs[0])

    document.add_heading("Sentence Evidence Register", level=1)
    document.add_paragraph(
        "Populate one row per sentence. Word counts use tokens of letters or numbers, allowing internal apostrophes and hyphens.",
        style="Evidence Note",
    )
    register = document.add_table(rows=1, cols=6)
    register.style = "Table Grid"
    headers = ["ID", "Source cue", "Source words", "Approved words", "Delta", "Edit class / review link"]
    for cell, header in zip(register.rows[0].cells, headers):
        set_cell_text(cell, header, bold=True, size=8.5)
        shade(cell, "1F4E79")
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
    repeat_header(register.rows[0])
    for sentence_id, *_ in APPROVED:
        row = register.add_row()
        set_cell_text(row.cells[0], sentence_id, bold=True, size=8.5)
        for cell in row.cells[1:]:
            set_cell_text(cell, "", size=8.5)
        keep_row(row)

    document.add_heading("Copy Integrity Summary", level=1)
    summary = document.add_table(rows=4, cols=2)
    summary.style = "Table Grid"
    for row, metric in zip(summary.rows, ["Source total", "Approved total", "Net change", "Unapproved factual additions"]):
        set_cell_text(row.cells[0], metric, bold=True)
        set_cell_text(row.cells[1], "")
        shade(row.cells[0], "D9EAF7")
        keep_row(row)

    document.add_heading("Evidence Note", level=1)
    document.add_paragraph(
        "Add a 90–130 word evidence note here; then change the review state in the metadata table to Approved copy applied.",
        style="Evidence Note",
    )
    document.core_properties.title = "Café Profile Editorial Review Template"
    document.core_properties.subject = "Controlled Word copy-edit workflow"
    document.save(target)


def build_ticket(target: Path) -> None:
    document = Document()
    configure_document(document)
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    document.add_heading("Café Profile — Approved Copy Ticket", 0)
    document.add_paragraph(
        "Ticket CP-016. Apply only the five approved sentence replacements below. The approved text may reorganize source wording but may not add proper nouns, dates, numbers, prices, menu items, addresses, ratings, opening hours, or measurable service promises.",
        style="Evidence Note",
    )
    table = document.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["ID", "Approved replacement", "Edit class", "Comment anchor", "Required comment"]
    for cell, header in zip(table.rows[0].cells, headers):
        set_cell_text(cell, header, bold=True, size=8.2)
        shade(cell, "1F4E79")
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = RGBColor(255, 255, 255)
    repeat_header(table.rows[0])
    for sentence_id, approved, edit_class, anchor, comment in APPROVED:
        row = table.add_row()
        for cell, value in zip(row.cells, [sentence_id, approved, edit_class, anchor, comment]):
            set_cell_text(cell, value, bold=value == sentence_id, size=8.2)
        keep_row(row)
    document.add_heading("Completion controls", level=1)
    controls = [
        "Retain the template's controlled source exactly; it remains the fact authority.",
        "Use Word-native tracked deletion and insertion for every sentence pair and enable Track Changes.",
        "Anchor the three required comments to the exact approved-copy phrases shown in the ticket.",
        "Create bookmarks review_S1 through review_S5 around the approved replacements and link each register row back to its sentence.",
        "Complete the register, integrity summary, evidence note, headers, footers, section geometry, and native editable tables; do not flatten content to images.",
    ]
    for item in controls:
        document.add_paragraph(item, style="List Bullet")
    document.core_properties.title = "Café Profile Approved Copy Ticket CP-016"
    document.save(target)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    source = args.source.read_text(encoding="utf-8")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    build_template(source, args.output_dir / "cafe_profile_editorial_template.docx")
    build_ticket(args.output_dir / "cafe_profile_approved_copy_ticket.docx")


if __name__ == "__main__":
    main()
