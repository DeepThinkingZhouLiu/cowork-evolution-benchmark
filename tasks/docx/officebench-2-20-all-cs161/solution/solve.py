#!/usr/bin/env python3
"""Build the task-specific CS161 roster control book."""

from __future__ import annotations

import argparse
import re
import zipfile
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook


OUTPUT_NAME = "cs161_roster_control_book.docx"
TITLE = "CS161 Roster Control Book"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]


def add_field(paragraph, instruction: str, placeholder: str) -> None:
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = placeholder
    run.append(text)
    field.append(run)
    paragraph._p.append(field)


def add_internal_link(paragraph, anchor: str, label: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    hyperlink.set(qn("w:history"), "1")
    run = OxmlElement("w:r")
    props = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    props.extend([color, underline])
    text = OxmlElement("w:t")
    text.text = label
    run.extend([props, text])
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def bookmark_cell(cell, bookmark_id: int, name: str) -> None:
    paragraph = cell.paragraphs[0]
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def shade(cell, fill: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(node)


def set_cell_margins(cell, top: int = 70, start: int = 75, bottom: int = 70, end: int = 75) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = margins.find(qn(f"w:{name}"))
        if node is None:
            node = OxmlElement(f"w:{name}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def format_table(table, widths: list[float] | None = None, small: bool = False) -> None:
    table.style = "Table Grid"
    table.autofit = False
    for row_index, row in enumerate(table.rows):
        props = row._tr.get_or_add_trPr()
        if row_index == 0:
            header = OxmlElement("w:tblHeader")
            header.set(qn("w:val"), "true")
            props.append(header)
        else:
            props.append(OxmlElement("w:cantSplit"))
        for column_index, cell in enumerate(row.cells):
            if widths:
                cell.width = Inches(widths[column_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_index == 0:
                shade(cell, "1F4E78")
            elif row_index % 2 == 0:
                shade(cell, "EAF2F8")
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                for run in paragraph.runs:
                    run.font.name = "Liberation Sans"
                    run.font.size = Pt(7.2 if small else 8.5)
                    if row_index == 0:
                        run.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)


def add_page_number_footer(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run("Page ")
    add_field(paragraph, "PAGE", "1")
    paragraph.add_run(" of ")
    add_field(paragraph, "NUMPAGES", "1")
    for run in paragraph.runs:
        run.font.name = "Liberation Sans"
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(89, 89, 89)


def add_header(section) -> None:
    section.header.is_linked_to_previous = False
    paragraph = section.header.paragraphs[0]
    paragraph.text = TITLE
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in paragraph.runs:
        run.font.name = "Liberation Sans"
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(89, 89, 89)


def configure_styles(document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(9.5)
    for style_name, size, color in (
        ("Title", 24, RGBColor(31, 78, 120)),
        ("Heading 1", 17, RGBColor(31, 78, 120)),
        ("Heading 2", 12, RGBColor(47, 84, 116)),
    ):
        style = document.styles[style_name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.font.bold = True


def load_source(path: Path) -> tuple[list[str], list[dict]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook["Sheet1"]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() for value in rows[0]]
    records = []
    for source_row, values in enumerate(rows[1:], start=2):
        if any(value is not None for value in values):
            records.append({**dict(zip(headers, values)), "Source row": f"Sheet1!{source_row}"})
    return headers, records


def add_distribution_table(document, records: list[dict]) -> None:
    table = document.add_table(rows=1, cols=4)
    for cell, label in zip(table.rows[0].cells, ["Dimension", "Category", "Count", "Share"]):
        cell.text = label
    total = len(records)
    for dimension, key in (("Major", "Major"), ("Academic standing", "Academic Standing")):
        counts = Counter(str(record[key]) for record in records)
        for category in sorted(counts):
            row = table.add_row().cells
            values = [dimension, category, str(counts[category]), f"{counts[category] / total * 100:.1f}%"]
            for cell, value in zip(row, values):
                cell.text = value
    format_table(table, [1.45, 1.75, 0.8, 1.0])


def add_roster_table(document, records: list[dict]) -> None:
    headers = ["Source row", "Student ID", "Last name", "First name", "Academic standing", "Major", *ASSESSMENTS]
    table = document.add_table(rows=1, cols=len(headers))
    for cell, label in zip(table.rows[0].cells, headers):
        cell.text = label
    for bookmark_id, record in enumerate(records, start=100):
        row = table.add_row().cells
        values = [
            record["Source row"], record["Student ID"], record["Last Name"], record["First Name"],
            record["Academic Standing"], record["Major"], *[record[name] for name in ASSESSMENTS],
        ]
        for cell, value in zip(row, values):
            cell.text = str(value)
        bookmark_cell(row[1], bookmark_id, f"roster_{record['Student ID']}")
    format_table(table, [0.75, 0.7, 0.85, 0.85, 1.0, 0.48, 0.43, 0.43, 0.55, 0.43, 0.43], small=True)


def add_assessment_table(document, records: list[dict]) -> None:
    table = document.add_table(rows=1, cols=5)
    for cell, label in zip(table.rows[0].cells, ["Assessment", "Valid N", "Average", "Minimum", "Maximum"]):
        cell.text = label
    for assessment in ASSESSMENTS:
        values = [int(record[assessment]) for record in records if record[assessment] is not None]
        row = table.add_row().cells
        output = [assessment, len(values), f"{sum(values) / len(values):.2f}", min(values), max(values)]
        for cell, value in zip(row, output):
            cell.text = str(value)
    format_table(table, [1.35, 0.9, 1.05, 1.05, 1.05])


def add_quality_table(document, headers: list[str], records: list[dict]) -> None:
    ids = [str(record["Student ID"]) for record in records]
    scores = [record[name] for record in records for name in ASSESSMENTS]
    complete = sum(all(record.get(header) is not None for header in headers) for record in records)
    valid_ids = sum(bool(re.fullmatch(r"S\d{4}", value)) for value in ids)
    valid_scores = sum(isinstance(value, int) and 0 <= value <= 100 for value in scores)
    checks = [
        ("Complete source rows", "PASS" if complete == len(records) else "FAIL", f"{complete} of {len(records)} rows have all 10 fields"),
        ("Unique Student IDs", "PASS" if len(set(ids)) == len(ids) else "FAIL", f"{len(set(ids))} unique of {len(ids)} IDs"),
        ("Student ID format", "PASS" if valid_ids == len(ids) else "FAIL", f"{valid_ids} of {len(ids)} match S + 4 digits"),
        ("Integer score bounds", "PASS" if valid_scores == len(scores) else "FAIL", f"{valid_scores} of {len(scores)} scores are integers in 0–100"),
    ]
    table = document.add_table(rows=1, cols=3)
    for cell, label in zip(table.rows[0].cells, ["Control", "Result", "Evidence"]):
        cell.text = label
    for values in checks:
        row = table.add_row().cells
        for cell, value in zip(row, values):
            cell.text = value
    format_table(table, [1.7, 0.8, 3.4])


def add_index_table(document, records: list[dict]) -> None:
    table = document.add_table(rows=1, cols=4)
    for cell, label in zip(table.rows[0].cells, ["Name", "Student ID", "Source row", "Roster link"]):
        cell.text = label
    for record in sorted(records, key=lambda item: (str(item["Last Name"]), str(item["First Name"]))):
        row = table.add_row().cells
        row[0].text = f"{record['Last Name']}, {record['First Name']}"
        row[1].text = str(record["Student ID"])
        row[2].text = record["Source row"]
        row[3].text = ""
        add_internal_link(row[3].paragraphs[0], f"roster_{record['Student ID']}", "Return to roster row")
    format_table(table, [1.8, 1.0, 1.0, 1.7])


def save_deterministic_docx(document, target: Path) -> None:
    """Normalize ZIP metadata so independent Oracle replays are byte-stable."""
    raw = target.with_name(f".{target.name}.raw")
    document.save(raw)
    with zipfile.ZipFile(raw) as source, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as output:
        for original in source.infolist():
            info = zipfile.ZipInfo(original.filename, date_time=(2026, 9, 7, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o600 << 16
            output.writestr(info, source.read(original.filename))
    raw.unlink()


def build(input_dir: Path, output_dir: Path) -> Path:
    headers, records = load_source(input_dir / "CS161_Class_Roster.xlsx")
    if headers != ["Student ID", "Last Name", "First Name", "Academic Standing", "HW1", "HW2", "Midterm", "HW3", "Final", "Major"]:
        raise ValueError(f"Unexpected source headers: {headers}")
    if len(records) != 9:
        raise ValueError(f"Expected 9 source rows, got {len(records)}")

    document = Document()
    configure_styles(document)
    document.core_properties.title = TITLE
    document.core_properties.subject = "Source-order roster and assessment reconciliation"
    document.core_properties.author = "CS Department Records Office"

    first = document.sections[0]
    first.top_margin = Inches(0.7)
    first.bottom_margin = Inches(0.65)
    first.left_margin = Inches(0.75)
    first.right_margin = Inches(0.75)
    add_header(first)
    add_page_number_footer(first)

    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(TITLE)
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("Registrar source-of-truth review edition").italic = True
    document.add_paragraph("Source file: CS161_Class_Roster.xlsx")
    document.add_paragraph("Worksheet: Sheet1")
    document.add_paragraph("Source data range: Sheet1!A2:J10 (rows 2–10)")
    document.add_paragraph("Prepared from the frozen workbook only; no external policy assumptions are applied.")
    document.add_heading("Contents", level=1)
    toc = document.add_paragraph()
    add_field(toc, 'TOC \\o "1-2" \\h \\z \\u', "Update field to refresh contents")

    document.add_page_break()
    document.add_heading("Control Summary", level=1)
    document.add_paragraph(f"The source contains {len(records)} populated student rows. The distributions below reconcile exactly to that control total.")
    add_distribution_table(document, records)

    roster_section = document.add_section(WD_SECTION.NEW_PAGE)
    roster_section.orientation = WD_ORIENT.LANDSCAPE
    roster_section.page_width, roster_section.page_height = first.page_height, first.page_width
    roster_section.top_margin = Inches(0.55)
    roster_section.bottom_margin = Inches(0.55)
    roster_section.left_margin = Inches(0.45)
    roster_section.right_margin = Inches(0.45)
    add_header(roster_section)
    add_page_number_footer(roster_section)
    document.add_heading("Source-order Roster Register", level=1)
    document.add_paragraph("Register order is identical to Sheet1 rows 2–10. Source row locators are part of the control evidence.")
    add_roster_table(document, records)

    audit_section = document.add_section(WD_SECTION.NEW_PAGE)
    audit_section.orientation = WD_ORIENT.PORTRAIT
    audit_section.page_width = Inches(8.5)
    audit_section.page_height = Inches(11)
    audit_section.top_margin = Inches(0.65)
    audit_section.bottom_margin = Inches(0.65)
    audit_section.left_margin = Inches(0.75)
    audit_section.right_margin = Inches(0.75)
    add_header(audit_section)
    add_page_number_footer(audit_section)

    document.add_heading("Assessment Reconciliation", level=1)
    document.add_paragraph("Valid N, averages and ranges are recomputed separately for each source assessment column.")
    add_assessment_table(document, records)
    document.add_heading("Data-quality Controls", level=1)
    add_quality_table(document, headers, records)
    document.add_page_break()
    document.add_heading("Alphabetical Name Index", level=1)
    document.add_paragraph("Index order is Last name, then First name. Each link returns to the corresponding source-order roster row.")
    add_index_table(document, records)
    total = len(records)
    document.add_paragraph(f"Control total: {total} source rows = {total} roster rows = {total} index entries.", style="Heading 2")

    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / OUTPUT_NAME
    save_deterministic_docx(document, target)
    return target


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
