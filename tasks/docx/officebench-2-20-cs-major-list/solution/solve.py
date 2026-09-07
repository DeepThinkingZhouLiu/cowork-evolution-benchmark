from __future__ import annotations

import argparse
import hashlib
import os
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


OUTPUT_NAME = "cs_major_advising_docket.docx"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
INPUT_SHA256 = "c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7"
NAVY = "17324D"
BLUE = "2D6A8A"
PALE_BLUE = "EAF2F7"
PALE_GOLD = "FFF3D6"
PALE_RED = "FCE8E6"
LIGHT = "F4F7F9"
WHITE = "FFFFFF"
INK = RGBColor(23, 50, 77)


def load_register(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    if hashlib.sha256(path.read_bytes()).hexdigest() != INPUT_SHA256:
        raise ValueError("The roster input does not match the frozen OfficeBench snapshot")
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [str(value).strip() if value is not None else "" for value in next(ws.iter_rows(values_only=True))]
    rows: list[dict[str, object]] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        rows.append(record)
    wb.close()
    return headers, rows


def tier_for(final_score: int) -> tuple[int, str, str]:
    if final_score <= 5:
        return 1, "Priority 1", "0-5"
    if final_score <= 20:
        return 2, "Priority 2", "6-20"
    return 3, "Priority 3", "21-100"


def student_truth(record: dict[str, object], headers: list[str]) -> dict[str, object]:
    scores = {name: int(record[name]) for name in ASSESSMENTS}
    values = list(scores.values())
    final_score = scores["Final"]
    midterm_delta = final_score - scores["Midterm"]
    tier_number, tier, band = tier_for(final_score)
    if midterm_delta < 0:
        trend = "Decline"
    elif midterm_delta > 0:
        trend = "Increase"
    else:
        trend = "Flat"
    low_value = min(values)
    high_value = max(values)
    low_names = [name for name in ASSESSMENTS if scores[name] == low_value]
    high_names = [name for name in ASSESSMENTS if scores[name] == high_value]
    return {
        **record,
        "name": f"{record['First Name']} {record['Last Name']}",
        "scores": scores,
        "average": round(mean(values), 1),
        "lowest": f"{' / '.join(low_names)} ({low_value})",
        "highest": f"{' / '.join(high_names)} ({high_value})",
        "range": high_value - low_value,
        "delta": midterm_delta,
        "trend": trend,
        "tier_number": tier_number,
        "tier": tier,
        "band": band,
        "assessment_columns": {name: get_column_letter(headers.index(name) + 1) for name in ASSESSMENTS},
    }


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top: int = 70, start: int = 90, bottom: int = 70, end: int = 90) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def mark_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tag = OxmlElement("w:tblHeader")
    tag.set(qn("w:val"), "true")
    tr_pr.append(tag)


def prevent_row_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def set_table_borders(table, color: str = "B7C8D6", size: str = "6") -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)
        borders.append(node)


def format_table(table, widths: list[float], header_fill: str = NAVY, font_size: float = 8.0) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    for r_idx, row in enumerate(table.rows):
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        prevent_row_split(row)
        if r_idx == 0:
            mark_repeat_header(row)
        for c_idx, cell in enumerate(row.cells):
            cell.width = Inches(widths[c_idx])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if r_idx == 0:
                set_cell_shading(cell, header_fill)
            elif r_idx % 2 == 0:
                set_cell_shading(cell, LIGHT)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    run.font.name = "Liberation Sans"
                    run.font.size = Pt(font_size)
                    if r_idx == 0:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)


def set_text(cell, text: object, *, bold: bool = False, color: RGBColor | None = None, align=None):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    if color is not None:
        run.font.color.rgb = color
    return run


def add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def add_internal_link(paragraph, text: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    run_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), BLUE)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    run_pr.extend((color, underline))
    run.append(run_pr)
    value = OxmlElement("w:t")
    value.text = text
    run.append(value)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def add_page_field(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Page ")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field_run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    field_run.append(text)
    field.append(field_run)
    paragraph._p.append(field)


def add_page_break(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)
    section.header_distance = Inches(0.22)
    section.footer_distance = Inches(0.25)

    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(9)
    normal.font.color.rgb = RGBColor(44, 57, 69)
    normal.paragraph_format.space_after = Pt(4)
    for style_name, size, color in (("Title", 25, INK), ("Heading 1", 16, INK), ("Heading 2", 12, RGBColor(45, 106, 138))):
        style = document.styles[style_name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = color
        style.paragraph_format.space_before = Pt(3)
        style.paragraph_format.space_after = Pt(5)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header_run = header.add_run("CS MAJOR REVIEW  •  SOURCE-CONTROLLED")
    header_run.font.name = "Liberation Sans"
    header_run.font.size = Pt(8)
    header_run.font.bold = True
    header_run.font.color.rgb = RGBColor(45, 106, 138)
    add_page_field(section.footer.paragraphs[0])


def add_title_page_content(document: Document, queue: list[dict[str, object]]) -> None:
    title = document.add_paragraph(style="Title")
    title.add_run("CS Major\nAdvising Review Docket")
    subtitle = document.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(7)
    run = subtitle.add_run("FILTERED EVIDENCE • MECHANICAL TRIAGE • AUDITABLE SOURCE CELLS")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = RGBColor(45, 106, 138)

    callout = document.add_table(rows=1, cols=2)
    callout.autofit = False
    set_text(callout.cell(0, 0), "SOURCE REGISTER", bold=True, color=RGBColor(255, 255, 255))
    set_text(callout.cell(0, 1), "CS161_Class_Roster.xlsx • Sheet1 • 9 source rows • exact Major = CS")
    set_cell_shading(callout.cell(0, 0), NAVY)
    set_cell_shading(callout.cell(0, 1), PALE_BLUE)
    for cell, width in zip(callout.rows[0].cells, (1.25, 6.05)):
        cell.width = Inches(width)
        set_cell_margins(cell, 100, 110, 100, 110)
    document.add_paragraph()

    document.add_heading("Review Queue", level=1)
    note = document.add_paragraph("Queue order is rule-driven: tier -> Final ascending -> Student ID ascending. Negative values show a decline from Midterm to Final.")
    note.paragraph_format.space_after = Pt(5)

    headers = ["Order", "Student ID", "Name", "Standing", "Final", "Δ M->F", "Tier", "Source row", "Evidence"]
    table = document.add_table(rows=1, cols=len(headers))
    for idx, label in enumerate(headers):
        set_text(table.cell(0, idx), label)
    for order, student in enumerate(queue, start=1):
        cells = table.add_row().cells
        values = [order, student["Student ID"], student["name"], student["Academic Standing"], student["scores"]["Final"], f"{student['delta']:+d}"]
        for idx, value in enumerate(values):
            set_text(cells[idx], value, align=WD_ALIGN_PARAGRAPH.CENTER if idx in (0, 4, 5) else None)
        tier_run = set_text(cells[6], student["tier"], bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_shading(cells[6], PALE_RED if student["tier"] == "Priority 1" else PALE_GOLD)
        comment = document.add_comment(
            tier_run,
            text=f"Tier rule: Final {student['band']} -> {student['tier']}. Observed Final: {student['scores']['Final']}.",
            author="Academic Review Control",
            initials="ARC",
        )
        comment._element.set(qn("w:date"), "2024-01-01T00:00:00Z")
        set_text(cells[7], f"Sheet1!{student['_row']}", align=WD_ALIGN_PARAGRAPH.CENTER)
        cells[8].text = ""
        link_paragraph = cells[8].paragraphs[0]
        link_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_internal_link(link_paragraph, "Open evidence record", f"student_{student['Student ID']}")
    format_table(table, [0.42, 0.68, 1.10, 0.80, 0.43, 0.58, 0.72, 0.72, 1.12], font_size=7.1)

    document.add_heading("Decision Rule Card", level=2)
    rules = document.add_table(rows=1, cols=3)
    for idx, label in enumerate(("Final band", "Review tier", "Interpretation")):
        set_text(rules.cell(0, idx), label)
    for band, tier, interpretation in (
        ("0-5", "Priority 1", "First review band"),
        ("6-20", "Priority 2", "Second review band"),
        ("21-100", "Priority 3", "Third review band"),
    ):
        row = rules.add_row().cells
        for idx, value in enumerate((band, tier, interpretation)):
            set_text(row[idx], value)
    format_table(rules, [1.55, 1.55, 4.15], font_size=8)


def add_profile(document: Document, student: dict[str, object], bookmark_id: int) -> None:
    add_page_break(document)
    heading = document.add_heading(f"Evidence Record • {student['Student ID']}", level=1)
    add_bookmark(heading, f"student_{student['Student ID']}", bookmark_id)
    intro = document.add_paragraph()
    lead = intro.add_run(f"{student['name']}  |  {student['Academic Standing']}  |  {student['tier']}")
    lead.bold = True
    lead.font.color.rgb = RGBColor(45, 106, 138)

    document.add_heading("Identity & Filter Evidence", level=2)
    identity = document.add_table(rows=1, cols=3)
    for idx, label in enumerate(("Field", "Value", "Source")):
        set_text(identity.cell(0, idx), label)
    identity_values = (
        ("Student ID", student["Student ID"], f"Sheet1!A{student['_row']}"),
        ("Name", student["name"], f"Sheet1!C{student['_row']} + B{student['_row']}"),
        ("Academic Standing", student["Academic Standing"], f"Sheet1!D{student['_row']}"),
        ("Major filter", student["Major"], f"Sheet1!J{student['_row']}"),
    )
    for values in identity_values:
        cells = identity.add_row().cells
        for idx, value in enumerate(values):
            set_text(cells[idx], value)
    format_table(identity, [1.8, 3.1, 2.35], font_size=8)

    document.add_heading("Assessment Evidence", level=2)
    scores = document.add_table(rows=1, cols=3)
    for idx, label in enumerate(("Assessment", "Score", "Source cell")):
        set_text(scores.cell(0, idx), label)
    for assessment in ASSESSMENTS:
        cells = scores.add_row().cells
        values = (assessment, student["scores"][assessment], f"Sheet1!{student['assessment_columns'][assessment]}{student['_row']}")
        for idx, value in enumerate(values):
            set_text(cells[idx], value, align=WD_ALIGN_PARAGRAPH.CENTER if idx == 1 else None)
    format_table(scores, [2.25, 1.25, 3.75], font_size=8)

    document.add_heading("Derived Checks", level=2)
    derived = document.add_table(rows=1, cols=3)
    for idx, label in enumerate(("Derived check", "Result", "Calculation basis")):
        set_text(derived.cell(0, idx), label)
    score_text = " + ".join(str(student["scores"][name]) for name in ASSESSMENTS)
    derived_values = (
        ("Five-assessment average", f"{student['average']:.1f}", f"({score_text}) / 5"),
        ("Lowest assessment", student["lowest"], "minimum of five scores"),
        ("Highest assessment", student["highest"], "maximum of five scores"),
        ("Score range", student["range"], "highest - lowest"),
        ("Midterm -> Final", f"{student['delta']:+d}", "Final - Midterm"),
        ("Trend", student["trend"], "sign of Midterm -> Final"),
        ("Review tier", student["tier"], f"Final {student['scores']['Final']} is in band {student['band']}"),
    )
    for values in derived_values:
        cells = derived.add_row().cells
        for idx, value in enumerate(values):
            set_text(cells[idx], value)
    format_table(derived, [2.0, 1.8, 3.45], font_size=7.8)


def add_appendix(document: Document, rows: list[dict[str, object]], selected: list[dict[str, object]]) -> None:
    add_page_break(document)
    document.add_heading("Cohort Comparison Appendix", level=1)
    paragraph = document.add_paragraph("Aggregate-only comparison. Non-CS identities and individual scores are intentionally suppressed.")
    paragraph.paragraph_format.space_after = Pt(6)

    selected_ids = {student["Student ID"] for student in selected}
    cs_rows = [row for row in rows if row["Student ID"] in selected_ids]
    non_cs_rows = [row for row in rows if row["Student ID"] not in selected_ids]
    table = document.add_table(rows=1, cols=6)
    headers = ["Assessment", "CS N", "CS avg", "Non-CS N", "Non-CS avg", "CS - non-CS"]
    for idx, label in enumerate(headers):
        set_text(table.cell(0, idx), label)
    for assessment in ASSESSMENTS:
        cs_average = mean(float(row[assessment]) for row in cs_rows)
        non_cs_average = mean(float(row[assessment]) for row in non_cs_rows)
        values = (assessment, len(cs_rows), f"{cs_average:.1f}", len(non_cs_rows), f"{non_cs_average:.1f}", f"{cs_average - non_cs_average:+.1f}")
        cells = table.add_row().cells
        for idx, value in enumerate(values):
            set_text(cells[idx], value, align=WD_ALIGN_PARAGRAPH.CENTER if idx else None)
    format_table(table, [1.55, 0.85, 1.1, 1.0, 1.25, 1.5], font_size=8)

    document.add_heading("Control Summary", level=1)
    controls = document.add_table(rows=1, cols=3)
    for idx, label in enumerate(("Control", "Result", "Evidence")):
        set_text(controls.cell(0, idx), label)
    source_rows = ", ".join(f"Sheet1!{student['_row']}" for student in sorted(selected, key=lambda item: int(item["_row"])))
    values = (
        ("Source row count", len(rows), "Sheet1 data rows 2-10"),
        ("CS selected count", len(selected), "Major exactly equals CS"),
        ("Non-CS excluded count", len(rows) - len(selected), "No individual disclosure in this docket"),
        ("Selected source rows", source_rows, "Complete filter trace"),
        ("Tier rules", "0-5 -> P1; 6-20 -> P2; 21-100 -> P3", "Applied to Final"),
        ("Ranking rule", "Tier -> Final asc -> Student ID asc", "Deterministic queue order"),
        ("Reconciliation", f"{len(selected)} queue rows = {len(selected)} evidence records", "Counts match"),
    )
    for row_values in values:
        cells = controls.add_row().cells
        for idx, value in enumerate(row_values):
            set_text(cells[idx], value)
    format_table(controls, [1.75, 3.25, 2.25], font_size=7.8)

    seal = document.add_paragraph()
    seal.paragraph_format.space_before = Pt(6)
    run = seal.add_run("CONTROL RESULT  •  FILTER COMPLETE  •  QUEUE RECONCILED  •  SOURCE CELLS TRACEABLE")
    run.bold = True
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(45, 106, 138)


def normalize_docx(path: Path) -> None:
    fixed_time = (2024, 1, 1, 0, 0, 0)
    with zipfile.ZipFile(path, "r") as source:
        entries = [(info.filename, source.read(info.filename), info.compress_type) for info in source.infolist()]
    fd, temp_name = tempfile.mkstemp(prefix="normalized-", suffix=".docx", dir=path.parent)
    os.close(fd)
    temp = Path(temp_name)
    try:
        with zipfile.ZipFile(temp, "w") as target:
            for name, data, compression in sorted(entries, key=lambda item: item[0]):
                info = zipfile.ZipInfo(name, date_time=fixed_time)
                info.compress_type = compression
                info.external_attr = 0o600 << 16
                target.writestr(info, data)
        temp.replace(path)
    finally:
        if temp.exists():
            temp.unlink()


def build(input_dir: Path, output_dir: Path) -> Path:
    headers, rows = load_register(input_dir / "CS161_Class_Roster.xlsx")
    selected = [student_truth(record, headers) for record in rows if str(record["Major"]).strip() == "CS"]
    queue = sorted(selected, key=lambda item: (int(item["tier_number"]), int(item["scores"]["Final"]), str(item["Student ID"])))
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME

    document = Document()
    document.core_properties.title = "CS Major Advising Review Docket"
    document.core_properties.subject = "Source-controlled CS-major review evidence"
    document.core_properties.author = "CoworkEvoBench reference implementation"
    document.core_properties.created = datetime(2024, 1, 1, tzinfo=timezone.utc)
    document.core_properties.modified = datetime(2024, 1, 1, tzinfo=timezone.utc)
    configure_document(document)
    add_title_page_content(document, queue)
    for bookmark_id, student in enumerate(queue, start=100):
        add_profile(document, student, bookmark_id)
    add_appendix(document, rows, selected)
    document.save(output)
    normalize_docx(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
