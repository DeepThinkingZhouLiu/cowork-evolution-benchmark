from __future__ import annotations

import argparse
import hashlib
import os
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean, median

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


OUTPUT_NAME = "perfect_final_verification_packet.docx"
INPUT_SHA256 = "c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7"
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
NAVY = "17324D"
BLUE = "2D6A8A"
LIGHT = "F2F6F8"
GREEN = "E4F2E9"
GOLD = "FFF3D6"


def load_rows(path: Path) -> tuple[list[str], list[dict[str, object]]]:
    if hashlib.sha256(path.read_bytes()).hexdigest() != INPUT_SHA256:
        raise ValueError("Roster does not match the frozen OfficeBench input")
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    headers = [str(value).strip() if value is not None else "" for value in next(ws.iter_rows(values_only=True))]
    rows = []
    for row_number, values in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value is not None for value in values):
            continue
        record = dict(zip(headers, values))
        record["_row"] = row_number
        rows.append(record)
    wb.close()
    return headers, rows


def add_repeat_header(row) -> None:
    prop = row._tr.get_or_add_trPr()
    prop.append(OxmlElement("w:tblHeader"))


def prevent_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def shade(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    tc_pr.append(node)


def margins(cell, value: int = 70) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for name in ("top", "start", "bottom", "end"):
        node = OxmlElement(f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)


def cell_text(cell, value: object, *, bold: bool = False, color: RGBColor | None = None, align=None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    if align is not None:
        paragraph.alignment = align
    run = paragraph.add_run(str(value))
    run.font.name = "Liberation Sans"
    run.font.size = Pt(7.5)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def border_table(table) -> None:
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "6")
        node.set(qn("w:color"), "B7C8D6")
        borders.append(node)
    table._tbl.tblPr.append(borders)


def make_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], font_size: float = 7.5):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        cell_text(table.cell(0, index), header, bold=True, color=RGBColor(255, 255, 255))
        shade(table.cell(0, index), NAVY)
    add_repeat_header(table.rows[0])
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell_text(cells[index], value)
    border_table(table)
    for row_index, row in enumerate(table.rows):
        prevent_split(row)
        for column_index, cell in enumerate(row.cells):
            cell.width = Inches(widths[column_index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)
            if row_index and row_index % 2 == 0:
                shade(cell, LIGHT)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return table


def add_bookmark(paragraph, name: str, bookmark_id: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def replace_with_link(cell, text: str, anchor: str) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
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


def add_equation(document: Document, expression: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(5)
    paragraph.paragraph_format.space_after = Pt(3)
    label = paragraph.add_run("Native verification equation  ")
    label.bold = True
    label.font.color.rgb = RGBColor(45, 106, 138)
    math_para = OxmlElement("m:oMathPara")
    math = OxmlElement("m:oMath")
    math_run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = expression
    math_run.append(text)
    math.append(math_run)
    math_para.append(math)
    paragraph._p.append(math_para)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Page ")
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
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)
    section.header_distance = Inches(0.2)
    section.footer_distance = Inches(0.2)
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(9)
    normal.font.color.rgb = RGBColor(42, 54, 66)
    normal.paragraph_format.space_after = Pt(3)
    for name, size in (("Title", 24), ("Heading 1", 16), ("Heading 2", 11)):
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(23, 50, 77) if name != "Heading 2" else RGBColor(45, 106, 138)
        style.paragraph_format.space_before = Pt(3)
        style.paragraph_format.space_after = Pt(4)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run("FINAL=100  •  VERIFICATION PACKET")
    run.font.name = "Liberation Sans"
    run.font.size = Pt(8)
    run.font.bold = True
    run.font.color.rgb = RGBColor(45, 106, 138)
    add_page_number(section.footer.paragraphs[0])


def add_population_page(document: Document, rows: list[dict[str, object]], selected: list[dict[str, object]]) -> None:
    document.add_heading("Perfect Final Score\nVerification Packet", level=0)
    subtitle = document.add_paragraph("EXACT THRESHOLD • COMPLETE POPULATION LEDGER • CELL-LEVEL TRACEABILITY")
    subtitle.runs[0].font.bold = True
    subtitle.runs[0].font.color.rgb = RGBColor(45, 106, 138)
    callout = make_table(document, ["SOURCE", "THRESHOLD"], [["CS161_Class_Roster.xlsx • Sheet1", "Selected only when Final = 100"]], [3.7, 3.6], 8.5)
    shade(callout.cell(1, 0), "EAF2F7")
    shade(callout.cell(1, 1), GREEN)

    finals = [int(row["Final"]) for row in rows]
    perfect = [value for value in finals if value == 100]
    below = [value for value in finals if value < 100]
    document.add_heading("Cohort Control", level=2)
    controls = [
        ["Source rows", len(rows), "Numeric Final", len(finals)],
        ["Exact 100", len(perfect), "Below 100", len(below)],
        ["Above 100", sum(value > 100 for value in finals), "Missing / non-numeric", 0],
        ["Perfect rate", f"{len(perfect) / len(rows) * 100:.1f}%", "Final mean", f"{mean(finals):.1f}"],
        ["Final median", f"{median(finals):.1f}", "Next-highest", max(below)],
    ]
    make_table(document, ["Control", "Result", "Control", "Result"], controls, [1.65, 1.2, 1.9, 2.55], 7.4)

    document.add_heading("Population Decision Ledger", level=1)
    ledger_rows = []
    for row in rows:
        final = int(row["Final"])
        decision = "Selected" if final == 100 else "Excluded"
        reason = "Final = 100" if final == 100 else f"Final below 100 by {100 - final}"
        ledger_rows.append([f"Sheet1!{row['_row']}", row["Student ID"], f"{row['First Name']} {row['Last Name']}", final, f"Sheet1!I{row['_row']}", decision, reason])
    ledger = make_table(document, ["Source row", "Student ID", "Name", "Final", "Final cell", "Decision", "Reason"], ledger_rows, [0.72, 0.72, 1.18, 0.48, 0.72, 0.70, 2.62], 6.8)
    selected_ids = {str(row["Student ID"]) for row in selected}
    for table_row in ledger.rows[1:]:
        if table_row.cells[1].text in selected_ids:
            replace_with_link(table_row.cells[5], "Selected", f"certificate_{table_row.cells[1].text}")
            shade(table_row.cells[5], GREEN)


def certificate_truth(row: dict[str, object], headers: list[str], all_rows: list[dict[str, object]]) -> dict[str, object]:
    scores = {name: int(row[name]) for name in ASSESSMENTS}
    pre = mean(scores[name] for name in ASSESSMENTS[:-1])
    overall = mean(scores.values())
    minimum = min(scores.values())
    maximum = max(scores.values())
    final_rank = 1 + sum(int(item["Final"]) > scores["Final"] for item in all_rows)
    return {
        **row,
        "name": f"{row['First Name']} {row['Last Name']}",
        "scores": scores,
        "pre_average": pre,
        "overall_average": overall,
        "lowest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == minimum)} ({minimum})",
        "highest": f"{' / '.join(name for name in ASSESSMENTS if scores[name] == maximum)} ({maximum})",
        "range": maximum - minimum,
        "rank": f"{final_rank} of {len(all_rows)} (tie)",
        "lift": scores["Final"] - pre,
        "columns": {name: get_column_letter(headers.index(name) + 1) for name in ASSESSMENTS},
    }


def add_certificate(document: Document, truth: dict[str, object], bookmark_id: int) -> None:
    page_break(document)
    heading = document.add_heading(f"Verification Certificate • {truth['Student ID']}", level=1)
    add_bookmark(heading, f"certificate_{truth['Student ID']}", bookmark_id)
    line = document.add_paragraph(f"{truth['name']}  |  Final {truth['scores']['Final']}  |  Exact-match certificate")
    line.runs[0].font.bold = True
    line.runs[0].font.color.rgb = RGBColor(45, 106, 138)

    document.add_heading("Identity Evidence", level=2)
    identity = [
        ["Student ID", truth["Student ID"], f"Sheet1!A{truth['_row']}"],
        ["Name", truth["name"], f"Sheet1!C{truth['_row']} + B{truth['_row']}"],
        ["Academic Standing", truth["Academic Standing"], f"Sheet1!D{truth['_row']}"],
        ["Major", truth["Major"], f"Sheet1!J{truth['_row']}"],
    ]
    make_table(document, ["Field", "Value", "Source cell"], identity, [1.8, 3.0, 2.5], 8)

    document.add_heading("Assessment Evidence", level=2)
    assessment = [[name, truth["scores"][name], f"Sheet1!{truth['columns'][name]}{truth['_row']}"] for name in ASSESSMENTS]
    make_table(document, ["Assessment", "Score", "Source cell"], assessment, [2.1, 1.2, 4.0], 8)

    document.add_heading("Verification Calculations", level=2)
    calculations = [
        ["Pre-Final average", f"{truth['pre_average']:.2f}", "mean of HW1, HW2, Midterm, HW3"],
        ["Five-assessment average", f"{truth['overall_average']:.2f}", "mean of all five scores"],
        ["Lowest assessment", truth["lowest"], "minimum of five scores"],
        ["Highest assessment", truth["highest"], "maximum of five scores"],
        ["Score range", truth["range"], "highest - lowest"],
        ["Final rank", truth["rank"], "competition rank; ties share rank"],
        ["Exact predicate", "TRUE: Final = 100", f"Sheet1!I{truth['_row']}"],
        ["Final lift", f"+{truth['lift']:.2f}", "Final - pre-Final average"],
    ]
    make_table(document, ["Check", "Result", "Basis"], calculations, [2.0, 2.1, 3.2], 7.4)
    add_equation(document, f"100 - {truth['pre_average']:.2f} = {truth['lift']:.2f}")


def add_closure_page(document: Document, rows: list[dict[str, object]], selected: list[dict[str, object]]) -> None:
    page_break(document)
    document.add_heading("Threshold Boundary & Closure", level=1)
    document.add_paragraph("The nearest excluded records prove that equality, not proximity, controls selection.")
    excluded = sorted((row for row in rows if int(row["Final"]) < 100), key=lambda row: (-int(row["Final"]), str(row["Student ID"])))[:3]
    boundary = [[f"Sheet1!{row['_row']}", row["Student ID"], f"{row['First Name']} {row['Last Name']}", int(row["Final"]), 100 - int(row["Final"]), f"Sheet1!I{row['_row']}", "Excluded"] for row in excluded]
    make_table(document, ["Source row", "Student ID", "Name", "Final", "Gap", "Final cell", "Decision"], boundary, [0.8, 0.8, 1.35, 0.55, 0.5, 0.8, 1.0], 7.2)

    document.add_heading("Final Distribution", level=2)
    finals = [int(row["Final"]) for row in rows]
    bands = [("0-19", 0, 19), ("20-39", 20, 39), ("40-59", 40, 59), ("60-79", 60, 79), ("80-99", 80, 99), ("100", 100, 100)]
    distribution = []
    for label, lower, upper in bands:
        count = sum(lower <= value <= upper for value in finals)
        distribution.append([label, count, f"{count / len(finals) * 100:.1f}%"])
    make_table(document, ["Final band", "Count", "Share"], distribution, [2.6, 1.2, 3.5], 8)

    document.add_heading("Closure Controls", level=2)
    selected_cells = ", ".join(f"Sheet1!I{row['_row']}" for row in selected)
    closure = [
        ["Population", f"{len(rows)} = {len(selected)} Selected + {len(rows) - len(selected)} Excluded", "Ledger reconciles"],
        ["Certificates", f"{len(selected)} = {len(selected)} Selected", "One certificate per selected row"],
        ["Selected Final cells", selected_cells, "Both cells equal 100"],
        ["Native equations", str(len(selected)), "One Final-lift equation per certificate"],
        ["Above threshold", "0", "No Final above 100"],
        ["Missing / non-numeric", "0", "All nine Final values numeric"],
    ]
    make_table(document, ["Control", "Result", "Evidence"], closure, [1.7, 3.2, 2.4], 7.6)
    seal = document.add_paragraph("CLOSURE  •  9 DECISIONS  •  2 CERTIFICATES  •  2 SOURCE CELLS  •  2 NATIVE EQUATIONS")
    seal.runs[0].font.bold = True
    seal.runs[0].font.color.rgb = RGBColor(45, 106, 138)


def normalize(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = [(item.filename, archive.read(item.filename), item.compress_type) for item in archive.infolist()]
    fd, temporary_name = tempfile.mkstemp(prefix="normalized-", suffix=".docx", dir=path.parent)
    os.close(fd)
    temporary = Path(temporary_name)
    try:
        with zipfile.ZipFile(temporary, "w") as archive:
            for name, data, compression in sorted(entries):
                info = zipfile.ZipInfo(name, date_time=(2024, 1, 1, 0, 0, 0))
                info.compress_type = compression
                info.external_attr = 0o600 << 16
                archive.writestr(info, data)
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()


def build(input_dir: Path, output_dir: Path) -> Path:
    headers, rows = load_rows(input_dir / "CS161_Class_Roster.xlsx")
    selected = [row for row in rows if isinstance(row["Final"], (int, float)) and int(row["Final"]) == 100]
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / OUTPUT_NAME
    document = Document()
    document.core_properties.title = "Perfect Final Score Verification Packet"
    document.core_properties.subject = "Final=100 exception verification"
    document.core_properties.author = "CoworkEvoBench reference implementation"
    document.core_properties.created = datetime(2024, 1, 1, tzinfo=timezone.utc)
    document.core_properties.modified = datetime(2024, 1, 1, tzinfo=timezone.utc)
    configure(document)
    add_population_page(document, rows, selected)
    truths = [certificate_truth(row, headers, rows) for row in selected]
    for bookmark_id, truth in enumerate(truths, start=200):
        add_certificate(document, truth, bookmark_id)
    add_closure_page(document, rows, selected)
    document.save(output)
    normalize(output)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    build(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
