from __future__ import annotations

import argparse
import hashlib
import os
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook


OUTPUT_NAME = "bill_course_history_audit.docx"
INPUT_SHA256 = "d7764b78e2931fbcda87f2a45ef140693c3d6977e0235b264aa868cf949bab8a"
GRADE_POINTS = {"A": 4.0, "A-": 3.7, "B+": 3.3, "B": 3.0}
NAVY = "183A50"
TEAL = "247A7A"
LIGHT = "F0F6F6"
GREEN = "E5F2EA"
GOLD = "FFF1CC"


def load_rows(path: Path) -> list[dict[str, object]]:
    if hashlib.sha256(path.read_bytes()).hexdigest() != INPUT_SHA256:
        raise ValueError("Bill course list does not match the frozen OfficeBench input")
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [str(value).strip() for value in next(ws.iter_rows(values_only=True))]
    rows: list[dict[str, object]] = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        record["_term"] = f"{record['Course Term']} {int(record['Year'])}"
        record["_points"] = GRADE_POINTS[str(record["Course Grade"])]
        subject, number = str(record["Course"]).split()
        record["_subject"] = subject
        record["_number"] = int(number)
        rows.append(record)
    wb.close()
    return rows


def summarize_terms(rows: list[dict[str, object]]) -> list[list[object]]:
    order: list[str] = []
    grouped: dict[str, list[dict[str, object]]] = defaultdict(list)
    for row in rows:
        term = str(row["_term"])
        if term not in grouped:
            order.append(term)
        grouped[term].append(row)
    output = []
    cumulative_count = 0
    cumulative_points = 0.0
    for term in order:
        group = grouped[term]
        term_points = sum(float(row["_points"]) for row in group)
        cumulative_count += len(group)
        cumulative_points += term_points
        output.append([
            term,
            " · ".join(str(row["Course"]) for row in group),
            len(group),
            f"{term_points:.2f}",
            f"{term_points / len(group):.2f}",
            cumulative_count,
            f"{cumulative_points:.2f}",
            f"{cumulative_points / cumulative_count:.2f}",
        ])
    return output


def summarize_subjects(rows: list[dict[str, object]]) -> list[list[object]]:
    output = []
    for subject in ("CS", "MATH", "PHYS"):
        group = [row for row in rows if row["_subject"] == subject]
        points = sum(float(row["_points"]) for row in group)
        output.append([subject, len(group), " → ".join(str(row["Course"]) for row in group), f"{points:.2f}", f"{points / len(group):.2f}"])
    return output


def continuity_rows(rows: list[dict[str, object]]) -> list[list[object]]:
    output = []
    for subject in ("CS", "MATH", "PHYS"):
        group = sorted((row for row in rows if row["_subject"] == subject), key=lambda row: int(row["_number"]))
        prior = None
        for row in group:
            number = int(row["_number"])
            if prior is None:
                output.append([subject, row["Course"], number, "—", "—", "START", f"Sheet1!C{row['_row']}"])
            else:
                delta = number - prior
                output.append([subject, row["Course"], number, prior, f"+{delta}", "CONTIGUOUS +1" if delta == 1 else f"JUMP +{delta}", f"Sheet1!C{row['_row']}"])
            prior = number
    return output


def repeat_header(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def prevent_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def shade(cell, color: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(node)


def margins(cell, value: int = 55) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = tc_pr.first_child_found_in("w:tcMar")
    if node is None:
        node = OxmlElement("w:tcMar")
        tc_pr.append(node)
    for edge in ("top", "start", "bottom", "end"):
        item = OxmlElement(f"w:{edge}")
        item.set(qn("w:w"), str(value))
        item.set(qn("w:type"), "dxa")
        node.append(item)


def cell_text(cell, value: object, *, bold: bool = False, white: bool = False, size: float = 7.2) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(value))
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.bold = bold
    if white:
        run.font.color.rgb = RGBColor(255, 255, 255)


def make_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], size: float = 7.2):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        cell_text(table.cell(0, index), header, bold=True, white=True, size=size)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for row_index, values in enumerate(rows, start=1):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell_text(cells[index], value, size=size)
            if row_index % 2 == 0:
                shade(cells[index], LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "B5C9CC")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)
    return table


def add_content_control(cell, *, tag: str, control_id: int) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    alias = OxmlElement("w:alias")
    alias.set(qn("w:val"), f"Review status for {tag.replace('TermReview_', '').replace('_', ' ')}")
    tag_node = OxmlElement("w:tag")
    tag_node.set(qn("w:val"), tag)
    identifier = OxmlElement("w:id")
    identifier.set(qn("w:val"), str(control_id))
    rich_text = OxmlElement("w:richText")
    properties.extend((alias, tag_node, identifier, rich_text))
    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), TEAL)
    run_properties.append(color)
    run.append(run_properties)
    text = OxmlElement("w:t")
    text.text = "Pending review"
    run.append(text)
    content.append(run)
    sdt.extend((properties, content))
    paragraph._p.append(sdt)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Bill Course Audit  •  Page ")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field_run = OxmlElement("w:r")
    field_text = OxmlElement("w:t")
    field_text.text = "1"
    field_run.append(field_text)
    field.append(field_run)
    paragraph._p.append(field)


def page_break(document: Document) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.add_run().add_break(WD_BREAK.PAGE)


def configure(document: Document) -> None:
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.header_distance = Inches(0.18)
    section.footer_distance = Inches(0.18)
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(8.5)
    normal.font.color.rgb = RGBColor(43, 55, 61)
    normal.paragraph_format.space_after = Pt(2.5)
    for name, size in (("Title", 21), ("Heading 1", 15), ("Heading 2", 10.5)):
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(24, 58, 80) if name != "Heading 2" else RGBColor(36, 122, 122)
        style.paragraph_format.space_before = Pt(3)
        style.paragraph_format.space_after = Pt(3)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run("BILL  •  COURSE HISTORY AUDIT")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(36, 122, 122)
    add_page_number(section.footer.paragraphs[0])


def add_intro(document: Document, rows: list[dict[str, object]]) -> None:
    document.add_heading("Bill Course History\nAudit Dossier", level=0)
    strap = document.add_paragraph("SOURCE-ORDER LEDGER • UNWEIGHTED GPA • CELL-LEVEL TRACEABILITY")
    strap.runs[0].font.bold = True
    strap.runs[0].font.color.rgb = RGBColor(36, 122, 122)
    callout = make_table(document, ["Frozen source", "Scope", "Calculation rule"], [["Bill_course_lists.xlsx • Sheet1", "15 nonblank records • rows 2–16", "Each course counts once; no credit weighting"]], [2.6, 2.35, 2.55], 7.7)
    for cell in callout.rows[1].cells:
        shade(cell, GREEN)
    document.add_heading("Grade-point policy supplied for this audit", level=2)
    make_table(document, ["Recorded grade", "A", "A-", "B+", "B"], [["Grade points", "4.0", "3.7", "3.3", "3.0"]], [1.4, 1.2, 1.2, 1.2, 1.2], 7.6)
    document.add_heading("Chronological Source Ledger", level=1)
    ledger = []
    for row in rows:
        ledger.append([f"Sheet1!{row['_row']}", row["Course Term"], int(row["Year"]), row["Course"], row["Course Grade"], f"{float(row['_points']):.1f}", f"Sheet1!A{row['_row']}:D{row['_row']}"])
    make_table(document, ["Source row", "Term", "Year", "Course", "Grade", "Points", "Source cells"], ledger, [0.8, 0.85, 0.52, 0.86, 0.55, 0.56, 1.55], 6.75)


def add_term_page(document: Document, rows: list[dict[str, object]]) -> None:
    document.add_heading("Term Performance & Review Queue", level=0)
    document.add_paragraph("Term order follows first appearance in the worksheet. GPA values are unweighted arithmetic means of the supplied grade points; displayed values are rounded to two decimals.")
    make_table(document, ["Term", "Courses", "N", "Points", "Term GPA", "Cum. N", "Cum. points", "Cum. GPA"], summarize_terms(rows), [0.9, 2.15, 0.35, 0.58, 0.62, 0.48, 0.72, 0.62], 6.75)
    document.add_heading("Editable Reviewer Status Controls", level=1)
    document.add_paragraph("Each status below is an unlocked native Word rich-text content control. Replace “Pending review” after evidence review; keep the tag stable for workflow automation.")
    terms = [row[0] for row in summarize_terms(rows)]
    control = make_table(document, ["Term", "Review status", "Stable control tag"], [[term, "", f"TermReview_{term.replace(' ', '_')}"] for term in terms], [1.55, 2.15, 3.1], 7.6)
    for index, term in enumerate(terms, start=1):
        add_content_control(control.cell(index, 1), tag=f"TermReview_{term.replace(' ', '_')}", control_id=8100 + index)
        shade(control.cell(index, 1), GOLD)
    document.add_heading("Term Control Notes", level=2)
    for note in (
        "Every term contains exactly three source records.",
        "Cumulative totals progress 3 → 6 → 9 → 12 → 15 courses.",
        "Final cumulative points are 54.00 and final cumulative GPA is 3.60.",
        "No credits appear in the source; no credit weighting is inferred.",
    ):
        document.add_paragraph(note, style="List Bullet")


def add_subject_page(document: Document, rows: list[dict[str, object]]) -> None:
    document.add_heading("Subject Progression & Code Continuity", level=0)
    document.add_paragraph("Subject totals use the source rows assigned to each course prefix. Continuity classifies adjacent numeric course codes only; it does not assert prerequisites or degree requirements.")
    document.add_heading("Subject Reconciliation", level=1)
    make_table(document, ["Subject", "N", "Source-order course sequence", "Points", "GPA"], summarize_subjects(rows), [0.75, 0.45, 3.65, 0.85, 0.75], 7.2)
    document.add_heading("Course-Code Continuity Map", level=1)
    make_table(document, ["Subject", "Course", "No.", "Prior no.", "Delta", "Classification", "Course cell"], continuity_rows(rows), [0.62, 0.76, 0.48, 0.62, 0.5, 1.25, 1.12], 6.6)


def add_closure_page(document: Document, rows: list[dict[str, object]]) -> None:
    document.add_heading("Grade Distribution & Audit Closure", level=0)
    document.add_paragraph("The distribution and closure below reconcile independently calculated row, term, subject, grade, and continuity totals.")
    counts = Counter(str(row["Course Grade"]) for row in rows)
    distribution = [[grade, f"{GRADE_POINTS[grade]:.1f}", counts[grade], f"{counts[grade] / len(rows) * 100:.1f}%", f"{GRADE_POINTS[grade] * counts[grade]:.2f}"] for grade in ("A", "A-", "B+", "B")]
    make_table(document, ["Grade", "Points", "Count", "Share", "Point contribution"], distribution, [0.85, 0.85, 0.85, 1.0, 1.65], 7.6)
    document.add_heading("Independent Closure Controls", level=1)
    closure = [
        ["Source population", "15", "Rows 2–16; no blank records"],
        ["Term balance", "15 = 3 × 5", "Three records in each of five terms"],
        ["Subject balance", "15 = 5 + 5 + 5", "CS + MATH + PHYS"],
        ["Grade balance", "15 = 5 + 4 + 4 + 2", "A + A- + B+ + B"],
        ["Point balance", "54.00 = 19.00 + 17.60 + 17.40", "CS + MATH + PHYS"],
        ["Overall GPA", "3.60", "54.00 ÷ 15; unweighted"],
        ["Course-code edges", "12 = 11 contiguous + 1 jump", "Only MATH 152 → 250 is JUMP +98"],
        ["Exact duplicate rows", "0", "Term, Year, Course, Grade tuple"],
        ["Blank required cells", "0", "A2:D16"],
        ["Review controls", "5", "One unlocked rich-text control per term"],
    ]
    table = make_table(document, ["Control", "Result", "Evidence / rule"], closure, [1.45, 2.25, 3.25], 7.2)
    for cell in table.rows[6].cells:
        shade(cell, GREEN)
    document.add_heading("Use Boundary", level=2)
    document.add_paragraph("This dossier reports only the frozen course-history evidence and the explicit grade-point policy above. It does not infer credits, prerequisites, degree completion, transfer status, or institutional standing.")


def build(input_path: Path, output_path: Path) -> None:
    rows = load_rows(input_path)
    document = Document()
    configure(document)
    add_intro(document, rows)
    page_break(document)
    add_term_page(document, rows)
    page_break(document)
    add_subject_page(document, rows)
    page_break(document)
    add_closure_page(document, rows)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    core = document.core_properties
    core.title = "Bill Course History Audit Dossier"
    core.subject = "Source-order course history and GPA reconciliation"
    core.author = "CoworkEvoBench Oracle"
    core.keywords = "course audit, source traceability, term GPA, content controls"
    core.comments = "Derived only from the frozen OfficeBench workbook and the explicit unweighted grade-point policy."
    document.save(output_path)


def normalize_docx(path: Path) -> None:
    fixed = (1980, 1, 1, 0, 0, 0)
    with tempfile.TemporaryDirectory(prefix="bill-audit-normalize-") as temp_dir:
        temp = Path(temp_dir) / path.name
        with zipfile.ZipFile(path) as incoming, zipfile.ZipFile(temp, "w") as outgoing:
            for info in sorted(incoming.infolist(), key=lambda item: item.filename):
                replacement = zipfile.ZipInfo(info.filename, fixed)
                replacement.compress_type = zipfile.ZIP_DEFLATED
                replacement.external_attr = info.external_attr
                outgoing.writestr(replacement, incoming.read(info.filename))
        os.replace(temp, path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=Path(os.environ.get("INPUT_DIR", "/app/data/input_files")) / "Bill_course_lists.xlsx")
    parser.add_argument("--output", type=Path, default=Path(os.environ.get("OUTPUT_DIR", "/app/output")) / OUTPUT_NAME)
    args = parser.parse_args()
    build(args.input, args.output)
    normalize_docx(args.output)


if __name__ == "__main__":
    main()
