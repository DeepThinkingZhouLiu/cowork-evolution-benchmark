from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from lxml import etree


INPUT_NAME = "homework.jpg"
MASTER_PATH = Path("homework.docx")
REVIEW_PATH = Path("homework/homework.docx")
INPUT_SHA256 = "1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf"
NAVY = "183C56"
TEAL = "28778B"
LIGHT = "F1F5F7"
GREEN = "E2F0E8"
GOLD = "FFF1CB"
ROSE = "FCE7E5"


def load_truth() -> dict[str, object]:
    candidates = [Path(__file__).parents[1] / "adaptation/package-truth.json", Path(__file__).with_name("package-truth.json")]
    for path in candidates:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    raise FileNotFoundError("package truth is unavailable")


def canonical_payload(truth: dict[str, object]) -> dict[str, object]:
    return {key: truth[key] for key in ("header", "instructions", "section_heading", "questions", "pipeline")}


def canonical_digest(truth: dict[str, object]) -> str:
    payload = json.dumps(canonical_payload(truth), ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def shade(cell, color: str) -> None:
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), color)
    cell._tc.get_or_add_tcPr().append(node)


def margins(cell, value: int = 55) -> None:
    properties = cell._tc.get_or_add_tcPr()
    node = properties.first_child_found_in("w:tcMar")
    if node is None:
        node = OxmlElement("w:tcMar")
        properties.append(node)
    for edge in ("top", "start", "bottom", "end"):
        margin = OxmlElement(f"w:{edge}")
        margin.set(qn("w:w"), str(value))
        margin.set(qn("w:type"), "dxa")
        node.append(margin)


def repeat_header(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def prevent_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def set_font(run, *, size: float, bold: bool = False, italic: bool = False, color: RGBColor | None = None, name: str = "Liberation Serif") -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def cell_text(cell, value: object, *, size: float = 7.0, bold: bool = False, white: bool = False, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    item = cell.paragraphs[0]
    item.alignment = align
    item.paragraph_format.space_after = Pt(0)
    run = item.add_run(str(value))
    set_font(run, size=size, bold=bold, color=RGBColor(255, 255, 255) if white else None, name="Liberation Sans")
    return run


def make_table(document: Document, headers: list[str], rows: list[list[object]], widths: list[float], *, size: float = 7.0, center_after_first: bool = False):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        align = WD_ALIGN_PARAGRAPH.CENTER if center_after_first and index else WD_ALIGN_PARAGRAPH.LEFT
        cell_text(table.cell(0, index), header, size=size, bold=True, white=True, align=align)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for row_index, values in enumerate(rows, start=1):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            align = WD_ALIGN_PARAGRAPH.CENTER if center_after_first and index else WD_ALIGN_PARAGRAPH.LEFT
            cell_text(cells[index], value, size=size, align=align)
            if row_index % 2 == 0:
                shade(cells[index], LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "A9BDC8")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches(widths[index])
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            margins(cell)
    return table


def paragraph(document: Document, text: str = "", *, size: float = 8.3, bold: bool = False, italic: bool = False, align=WD_ALIGN_PARAGRAPH.LEFT, before: float = 0, after: float = 2, left: float = 0):
    item = document.add_paragraph()
    item.alignment = align
    item.paragraph_format.space_before = Pt(before)
    item.paragraph_format.space_after = Pt(after)
    item.paragraph_format.left_indent = Inches(left)
    if text:
        set_font(item.add_run(text), size=size, bold=bold, italic=italic)
    return item


def heading(document: Document, text: str, *, size: float = 16, level: int = 1, before: float = 4, after: float = 3):
    item = document.add_paragraph(style=f"Heading {level}")
    item.paragraph_format.space_before = Pt(before)
    item.paragraph_format.space_after = Pt(after)
    set_font(item.add_run(text), size=size, bold=True, color=RGBColor(24, 60, 86), name="Liberation Sans")
    return item


def page_break(document: Document) -> None:
    item = document.add_paragraph()
    item.paragraph_format.space_after = Pt(0)
    item.add_run().add_break(WD_BREAK.PAGE)


def add_bookmark(item, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    item._p.insert(0, start)
    item._p.append(end)


def add_internal_link(item, label: str, anchor: str) -> None:
    link = OxmlElement("w:hyperlink")
    link.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), TEAL)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend((color, underline))
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    item._p.append(link)


def add_page_number(item, prefix: str) -> None:
    item.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(item.add_run(prefix), size=7.2, name="Liberation Sans")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = "1"
    run.append(text)
    field.append(run)
    item._p.append(field)


def add_equation(document: Document, expression: str) -> None:
    item = document.add_paragraph()
    item.alignment = WD_ALIGN_PARAGRAPH.CENTER
    item.paragraph_format.space_after = Pt(2)
    math_para = OxmlElement("m:oMathPara")
    math = OxmlElement("m:oMath")
    run = OxmlElement("m:r")
    text = OxmlElement("m:t")
    text.text = expression
    run.append(text)
    math.append(run)
    math_para.append(math)
    item._p.append(math_para)


def add_content_control(cell, *, tag: str, alias_text: str, value: str, control_id: int) -> None:
    cell.text = ""
    item = cell.paragraphs[0]
    item.paragraph_format.space_after = Pt(0)
    sdt = OxmlElement("w:sdt")
    properties = OxmlElement("w:sdtPr")
    alias = OxmlElement("w:alias")
    alias.set(qn("w:val"), alias_text)
    tag_node = OxmlElement("w:tag")
    tag_node.set(qn("w:val"), tag)
    identifier = OxmlElement("w:id")
    identifier.set(qn("w:val"), str(control_id))
    rich = OxmlElement("w:richText")
    properties.extend((alias, tag_node, identifier, rich))
    content = OxmlElement("w:sdtContent")
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), TEAL)
    run_properties.append(color)
    run.append(run_properties)
    text = OxmlElement("w:t")
    text.text = value
    run.append(text)
    content.append(run)
    sdt.extend((properties, content))
    item._p.append(sdt)


def configure(document: Document, role: str) -> None:
    section = document.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.46)
    section.bottom_margin = Inches(0.46)
    section.left_margin = Inches(0.48)
    section.right_margin = Inches(0.48)
    section.header_distance = Inches(0.21)
    section.footer_distance = Inches(0.21)
    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    label = "CONTROLLED MASTER" if role == "master" else "QA WORKING COPY"
    set_font(header.add_run(f"CSE 240A  •  {label}  •  OCR-HW2-1185885B"), size=7.1, bold=True, color=RGBColor(40, 119, 139), name="Liberation Sans")
    add_page_number(section.footer.paragraphs[0], f"OCR-HW2-1185885B  •  {label}  •  Page ")
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Serif"
    normal.font.size = Pt(8.3)


def title_block(document: Document, title: str, subtitle: str) -> None:
    item = document.add_paragraph(style="Title")
    item.alignment = WD_ALIGN_PARAGRAPH.LEFT
    item.paragraph_format.space_after = Pt(2)
    set_font(item.add_run(title), size=23, bold=True, color=RGBColor(24, 60, 86), name="Liberation Sans")
    paragraph(document, subtitle, size=8.2, bold=True, after=7)


def navigation(document: Document, links: list[tuple[str, str]]) -> None:
    item = paragraph(document, "Open: ", size=8, bold=True, after=5)
    for index, (label, anchor) in enumerate(links):
        add_internal_link(item, label, anchor)
        if index + 1 != len(links):
            set_font(item.add_run("  •  "), size=8, name="Liberation Sans")


def expected_pipeline(truth: dict[str, object]) -> list[list[str]]:
    arrows = {
        ("lw R2, 0 (R5)", "5"): "W\n↓",
        ("add R3, R2, R2", "8"): "W\n↓",
        ("add R2, R2, R5", "9"): "M\n↓",
    }
    rows = []
    for row in truth["pipeline"]["rows"]:
        instruction = row["instruction"]
        rows.append([instruction, *[arrows.get((instruction, cycle), row["stages"].get(cycle, "")) for cycle in truth["pipeline"]["cycles"]]])
    return rows


def question(document: Document, item: dict[str, object]) -> None:
    output = paragraph(document, size=7.75, after=1.4, left=0.1)
    set_font(output.add_run(f"{item['number']}. "), size=7.75, bold=True)
    set_font(output.add_run(item["prompt"]), size=7.75)


def build_transcript(document: Document, truth: dict[str, object], bookmark_id: int) -> None:
    page_break(document)
    title = heading(document, "Canonical Source Payload", size=18)
    add_bookmark(title, "canonical_payload", bookmark_id)
    paragraph(document, "Normalized for editing; source words, punctuation, order, contradictions, matrix cells, and worked content remain unchanged.", size=7.6, italic=True, after=4)
    header = truth["header"]
    paragraph(document, header["title"], size=15.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, after=0)
    paragraph(document, header["subtitle"], size=10, align=WD_ALIGN_PARAGRAPH.CENTER, after=3)
    paragraph(document, header["identity"], size=9, align=WD_ALIGN_PARAGRAPH.CENTER, after=1)
    paragraph(document, header["due"], size=8.6, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
    paragraph(document, "Instructions", size=8.7, bold=True, after=1)
    for text in truth["instructions"]:
        item = document.add_paragraph(style="List Bullet")
        item.paragraph_format.space_after = Pt(0.4)
        item.paragraph_format.left_indent = Inches(0.2)
        item.paragraph_format.first_line_indent = Inches(-0.1)
        set_font(item.add_run(text), size=7.25)
    paragraph(document, truth["section_heading"], size=10.5, bold=True, before=1, after=1)
    question(document, truth["questions"][0])
    paragraph(document, truth["questions"][0]["note"], size=7.0, left=0.2, after=1.5)
    table = make_table(document, ["Instruction", *truth["pipeline"]["cycles"]], expected_pipeline(truth), [1.72, *([0.35] * 15)], size=5.9, center_after_first=True)
    for cell in table.rows[1].cells:
        shade(cell, GREEN)
    question(document, truth["questions"][1])
    paragraph(document, truth["questions"][1]["note"], size=7.0, left=0.2, after=1.5)
    question(document, truth["questions"][2])
    add_equation(document, truth["questions"][2]["equation"])
    paragraph(document, truth["questions"][2]["conclusion"], size=7.6, left=0.2, after=0)


def canonical_blocks(truth: dict[str, object]) -> list[list[str]]:
    blocks: list[tuple[str, str, str]] = [
        ("H1", "title", truth["header"]["title"]),
        ("H2", "subtitle", truth["header"]["subtitle"]),
        ("H3", "identity", truth["header"]["identity"]),
        ("H4", "due", truth["header"]["due"]),
        ("S0", "label", "Instructions"),
    ]
    blocks.extend((f"I{index}", "instruction", text) for index, text in enumerate(truth["instructions"], 1))
    blocks.append(("S1", "section", truth["section_heading"]))
    for item in truth["questions"]:
        blocks.append((f"Q{item['number']}", "prompt", f"{item['number']}. {item['prompt']}"))
        if "note" in item:
            blocks.append((f"N{item['number']}", "note", item["note"]))
        if "equation" in item:
            blocks.append(("E3", "equation", item["equation"]))
            blocks.append(("C3", "conclusion", item["conclusion"]))
    return [[key, role, str(len(text)), hashlib.sha256(text.encode("utf-8")).hexdigest()[:12], text[:58] + ("…" if len(text) > 58 else "")] for key, role, text in blocks]


def master_cover(document: Document, truth: dict[str, object], digest: str) -> None:
    title_block(document, "Controlled Digitization Master", "CANONICAL PAYLOAD • TWO-COPY GOVERNANCE • SOURCE-BOUND RELEASE")
    navigation(document, [("Canonical payload", "canonical_payload"), ("Payload register", "payload_register"), ("Master handoff", "master_handoff")])
    package = truth["package"]
    heading(document, "Package Control", size=13, level=2)
    rows = [
        ["Package ID", package["package_id"], "Stable across both copies"],
        ["Version", package["version"], "Controlled remake"],
        ["Source", "homework.jpg • 1700 × 2200 px", "One frozen visible page"],
        ["Source SHA256", truth["source"]["sha256"], "Exact runtime bytes"],
        ["Canonical payload SHA256", digest, "Exact source-visible data model"],
        ["Document role", package["master_role"], package["master_path"]],
        ["Release state", package["release_state"], "Conflicts remain disclosed"],
    ]
    table = make_table(document, ["Control", "Value", "Basis"], rows, [1.55, 3.65, 2.25], size=6.8)
    for cell in table.rows[5].cells:
        shade(cell, GREEN)
    heading(document, "Copy Role Matrix", size=13, level=2)
    make_table(document, ["Required path", "Role", "Role-specific layer", "Canonical payload"], [
        [package["master_path"], package["master_role"], "Register + release closure", digest[:16]],
        [package["review_path"], package["review_role"], "Comments + unlocked QA controls", digest[:16]],
    ], [1.45, 1.85, 2.75, 1.4], size=6.8)
    heading(document, "Release Boundary", size=13, level=2)
    paragraph(document, "The master publishes only source-visible content plus audit metadata. The paired review copy carries unresolved decisions. Neither copy invents the absent Question 2 diagram, answers, grades, or institutional policy.", size=8.1, after=0)


def build_master_register(document: Document, truth: dict[str, object]) -> None:
    page_break(document)
    title = heading(document, "Canonical Payload Register", size=18)
    add_bookmark(title, "payload_register", 7102)
    paragraph(document, "Every source-visible text block is identified independently of Word pagination. SHA-12 is computed from exact verified text before display whitespace normalization.", size=7.5, after=3)
    make_table(document, ["Block", "Role", "Chars", "SHA-12", "Verified locator"], canonical_blocks(truth), [0.48, 0.9, 0.48, 1.05, 4.54], size=6.1)
    heading(document, "Native Object Register", size=12.5, level=2, before=2)
    objects = [
        ["Pipeline matrix", "1", "Native 7 × 16 table", "Both copies"],
        ["Worked equation", "1", "Native OMML", "Both copies"],
        ["Source image", "0", "Not embedded or flattened", "Both copies"],
        ["Bookmarks / links", "3 / 3", "Internal navigation", "Each copy"],
        ["Review comments", "6", "Source-bound decisions", "Review copy only"],
        ["Review controls", "4", "Unlocked rich-text controls", "Review copy only"],
    ]
    make_table(document, ["Object", "Count", "Editability", "Scope"], objects, [1.5, 0.65, 3.25, 2.05], size=6.5)


def build_master_handoff(document: Document, truth: dict[str, object], digest: str) -> None:
    page_break(document)
    title = heading(document, "Master Handoff & Pair Integrity", size=18)
    add_bookmark(title, "master_handoff", 7103)
    heading(document, "Cross-Copy Contract", size=13, level=2)
    make_table(document, ["Checkpoint", "Master", "Review", "Required relation"], [
        ["Package ID", truth["package"]["package_id"], truth["package"]["package_id"], "Equal"],
        ["Source SHA256", INPUT_SHA256[:16], INPUT_SHA256[:16], "Equal; exact full property"],
        ["Canonical payload SHA256", digest[:16], digest[:16], "Equal; exact full property"],
        ["Canonical text blocks", "18", "18", "Exact and ordered"],
        ["Pipeline rows / cycles", "6 / 15", "6 / 15", "Exact cells and arrows"],
        ["Document role", "MASTER", "REVIEW", "Must differ"],
        ["Role-specific controls", "None", "4 controls + 6 comments", "Must differ"],
    ], [1.55, 1.65, 1.65, 2.6], size=6.7)
    heading(document, "Source Exceptions", size=13, level=2)
    make_table(document, ["ID", "Evidence", "Master treatment"], [
        ["D01", truth["review_issues"][0][2], truth["review_issues"][0][3]],
        ["D02", truth["review_issues"][1][2], truth["review_issues"][1][3]],
    ], [0.5, 3.8, 3.15], size=6.7)
    heading(document, "Master Closure", size=13, level=2)
    closure = [
        ["Required DOCX paths", "2", "Root master + nested review copy"],
        ["Byte-identical copies", "No", "Role layers intentionally differ"],
        ["Canonical payload synchronized", "Yes", digest],
        ["Embedded media", "0", "Editable Word objects only"],
        ["Source conflicts disclosed", "2", "No repair or unsupported inference"],
        ["Review queue items", "6", "Resolved only in working copy workflow"],
        ["External relationships", "0", "Offline-safe package"],
        ["Release authority", "Master", "Review copy cannot self-approve"],
    ]
    table = make_table(document, ["Control", "Result", "Evidence / rule"], closure, [2.0, 1.2, 4.25], size=6.6)
    for cell in table.rows[3].cells:
        shade(cell, GREEN)


def review_cover(document: Document, truth: dict[str, object], digest: str) -> None:
    title_block(document, "OCR Review Working Copy", "EDITABLE QA LAYER • SOURCE AMBIGUITIES • SYNCHRONIZED PAYLOAD")
    navigation(document, [("Canonical payload", "canonical_payload"), ("Review queue", "review_queue"), ("Reviewer handoff", "review_handoff")])
    heading(document, "Review Controls", size=13, level=2)
    controls = document.add_table(rows=1, cols=3)
    controls.alignment = WD_TABLE_ALIGNMENT.CENTER
    controls.autofit = False
    for index, label in enumerate(("Control", "Editable value", "Purpose")):
        cell_text(controls.cell(0, index), label, size=7, bold=True, white=True)
        shade(controls.cell(0, index), NAVY)
    repeat_header(controls.rows[0])
    purposes = ["Workflow stage", "Named reviewer", "Source treatment", "Publication gate"]
    for index, (tag, alias, value) in enumerate(truth["package"]["review_controls"]):
        cells = controls.add_row().cells
        cell_text(cells[0], tag, size=6.9)
        add_content_control(cells[1], tag=tag, alias_text=alias, value=value, control_id=8200 + index)
        cell_text(cells[2], purposes[index], size=6.9)
        if index % 2:
            for cell in cells:
                shade(cell, LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "A9BDC8")
        borders.append(border)
    controls._tbl.tblPr.append(borders)
    for row in controls.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches([1.55, 3.45, 2.45][index])
            margins(cell)
    heading(document, "Copy Role Matrix", size=13, level=2)
    make_table(document, ["This path", "Role", "Paired master", "Canonical SHA"], [[truth["package"]["review_path"], truth["package"]["review_role"], truth["package"]["master_path"], digest[:16]]], [2.0, 1.6, 1.55, 2.3], size=6.8)
    heading(document, "QA Boundary", size=13, level=2)
    paragraph(document, "Change only the four unlocked controls during review. Source-visible payload and issue evidence remain synchronized with the controlled master; a reviewer must not manufacture the absent diagram or silently normalize conflicts.", size=8.1, after=0)


def build_review_queue(document: Document, truth: dict[str, object]) -> None:
    page_break(document)
    title = heading(document, "Ambiguity Review Queue", size=18)
    add_bookmark(title, "review_queue", 7202)
    paragraph(document, "Each disposition is a native Word comment target and remains source-bound. Comments explain the control decision; they are not hidden answers.", size=7.6, after=4)
    table = document.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, label in enumerate(("ID", "Issue", "Source evidence", "Required disposition")):
        cell_text(table.cell(0, index), label, size=6.8, bold=True, white=True)
        shade(table.cell(0, index), NAVY)
    repeat_header(table.rows[0])
    for index, issue in enumerate(truth["review_issues"]):
        cells = table.add_row().cells
        cell_text(cells[0], issue[0], size=6.6, bold=True)
        cell_text(cells[1], issue[1], size=6.6)
        cell_text(cells[2], issue[2], size=6.6)
        target = cell_text(cells[3], issue[3], size=6.6)
        comment = document.add_comment(target, text=f"{issue[0]} control: {issue[3]}", author="OCR QA Control", initials="OQC")
        comment._element.set(qn("w:date"), "2026-09-08T00:00:00Z")
        if index in (0, 1):
            for cell in cells:
                shade(cell, GOLD)
        elif index % 2:
            for cell in cells:
                shade(cell, LIGHT)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "5")
        border.set(qn("w:color"), "A9BDC8")
        borders.append(border)
    table._tbl.tblPr.append(borders)
    for row in table.rows:
        prevent_split(row)
        for index, cell in enumerate(row.cells):
            cell.width = Inches([0.48, 1.55, 2.85, 2.57][index])
            margins(cell, 70)
    heading(document, "Comment Contract", size=13, level=2)
    make_table(document, ["Control", "Required value", "Reason"], [
        ["Comment author", "OCR QA Control / OQC", "Stable audit provenance"],
        ["Comment count", "6", "One per review issue"],
        ["Decision state", "Preserve / do not invent", "Source-first review"],
        ["Tracked approval", "None", "No self-certified release"],
    ], [1.45, 2.3, 3.7], size=6.8)


def build_review_handoff(document: Document, truth: dict[str, object], digest: str) -> None:
    page_break(document)
    title = heading(document, "Reviewer Handoff & Synchronization", size=18)
    add_bookmark(title, "review_handoff", 7203)
    heading(document, "Control Inventory", size=13, level=2)
    make_table(document, ["Tag", "Alias", "Initial value", "Lock state"], [[tag, alias, value, "Unlocked"] for tag, alias, value in truth["package"]["review_controls"]], [1.55, 1.45, 3.1, 1.35], size=6.8)
    heading(document, "Cross-Copy Checklist", size=13, level=2)
    checks = [
        ["Package ID", truth["package"]["package_id"], "Must match master property"],
        ["Source SHA256", INPUT_SHA256, "Must match master and runtime JPEG"],
        ["Canonical payload SHA256", digest, "Must match master property"],
        ["Canonical block count", "18", "Exact source-visible order"],
        ["Pipeline table", "6 rows × 15 cycles", "Exact cells and 3 down arrows"],
        ["Native equation", "1", "Exact timing relation"],
        ["Review comments", "6", "One source-bound decision each"],
        ["Review controls", "4", "Exact tags, aliases, initial values; unlocked"],
        ["Embedded source image", "0", "No flattened-page substitute"],
        ["External relationships", "0", "Offline-safe"],
    ]
    make_table(document, ["Checkpoint", "Expected", "Verification"], checks, [1.85, 3.0, 2.6], size=6.45)
    heading(document, "Reviewer Boundary", size=13, level=2)
    for text in (
        "Verify against the frozen source before changing a QA control.",
        "Do not edit the canonical transcript, pipeline cells, equation, or issue evidence.",
        "Do not mark the copy ready for release while D01 or D02 remains unresolved by source evidence.",
        "Return the reviewed working copy alongside the unchanged controlled master.",
    ):
        item = document.add_paragraph(style="List Bullet")
        item.paragraph_format.space_after = Pt(1)
        set_font(item.add_run(text), size=7.7)


def custom_properties_xml(truth: dict[str, object], *, role: str, copy_path: str, digest: str) -> bytes:
    custom_ns = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
    vt_ns = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    root = etree.Element(f"{{{custom_ns}}}Properties", nsmap={None: custom_ns, "vt": vt_ns})
    values = [
        ("PackageID", truth["package"]["package_id"]),
        ("PackageVersion", truth["package"]["version"]),
        ("SourceSHA256", INPUT_SHA256),
        ("CanonicalPayloadSHA256", digest),
        ("DocumentRole", role),
        ("CopyPath", copy_path),
        ("PairCount", "2"),
        ("ReleaseState", truth["package"]["release_state"]),
    ]
    for pid, (name, value) in enumerate(values, start=2):
        prop = etree.SubElement(root, f"{{{custom_ns}}}property")
        prop.set("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        prop.set("pid", str(pid))
        prop.set("name", name)
        child = etree.SubElement(prop, f"{{{vt_ns}}}lpwstr")
        child.text = str(value)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def inject_properties(path: Path, truth: dict[str, object], *, role: str, copy_path: str, digest: str) -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {info.filename: (info, archive.read(info.filename)) for info in archive.infolist()}
    content_types = etree.fromstring(entries["[Content_Types].xml"][1])
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    override = etree.SubElement(content_types, f"{{{ct_ns}}}Override")
    override.set("PartName", "/docProps/custom.xml")
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.custom-properties+xml")
    entries["[Content_Types].xml"] = (entries["[Content_Types].xml"][0], etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    relationships = etree.fromstring(entries["_rels/.rels"][1])
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    relation = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
    relation.set("Id", "rId99")
    relation.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
    relation.set("Target", "docProps/custom.xml")
    entries["_rels/.rels"] = (entries["_rels/.rels"][0], etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes"))
    entries["docProps/custom.xml"] = (None, custom_properties_xml(truth, role=role, copy_path=copy_path, digest=digest))
    with tempfile.TemporaryDirectory(prefix="controlled-package-") as temp_dir:
        target = Path(temp_dir) / path.name
        fixed = (1980, 1, 1, 0, 0, 0)
        with zipfile.ZipFile(target, "w") as outgoing:
            for name in sorted(entries):
                info, payload = entries[name]
                replacement = zipfile.ZipInfo(name, fixed)
                replacement.compress_type = zipfile.ZIP_DEFLATED
                replacement.external_attr = info.external_attr if info is not None else 0o600 << 16
                outgoing.writestr(replacement, payload)
        os.replace(target, path)


def finalize(document: Document, path: Path, truth: dict[str, object], *, role: str, digest: str) -> None:
    core = document.core_properties
    if role == "MASTER":
        core.title = "CSE 240A Homework 2 Controlled Digitization Master"
        core.subject = "Canonical editable source payload and two-copy release control"
        core.comments = "Master copy; QA decisions remain in the paired working copy."
    else:
        core.title = "CSE 240A Homework 2 OCR Review Working Copy"
        core.subject = "Editable source payload, comments, and review controls"
        core.comments = "Working copy; not self-authorized for release."
    core.author = "CoworkEvoBench Oracle"
    core.category = "Controlled OCR package"
    core.keywords = "CSE 240A, OCR, controlled copy, review comments, content controls"
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    inject_properties(path, truth, role=role, copy_path=str(MASTER_PATH if role == "MASTER" else REVIEW_PATH), digest=digest)


def build(input_path: Path, output_dir: Path) -> tuple[Path, Path]:
    if hashlib.sha256(input_path.read_bytes()).hexdigest() != INPUT_SHA256:
        raise ValueError("Input does not match the frozen OfficeBench scan")
    truth = load_truth()
    if truth["source"]["sha256"] != INPUT_SHA256:
        raise ValueError("Package truth and runtime source disagree")
    digest = canonical_digest(truth)

    master = Document()
    configure(master, "master")
    master_cover(master, truth, digest)
    build_transcript(master, truth, 7101)
    build_master_register(master, truth)
    build_master_handoff(master, truth, digest)

    review = Document()
    configure(review, "review")
    review_cover(review, truth, digest)
    build_transcript(review, truth, 7201)
    build_review_queue(review, truth)
    build_review_handoff(review, truth, digest)

    master_path = output_dir / MASTER_PATH
    review_path = output_dir / REVIEW_PATH
    finalize(master, master_path, truth, role="MASTER", digest=digest)
    finalize(review, review_path, truth, role="REVIEW", digest=digest)
    if master_path.read_bytes() == review_path.read_bytes():
        raise AssertionError("role-specific copies must not be byte-identical")
    return master_path, review_path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", type=Path, default=Path(os.environ.get("INPUT_DIR", "/app/data/input_files")) / INPUT_NAME)
    parser.add_argument("--output-dir", type=Path, default=Path(os.environ.get("OUTPUT_DIR", "/app/output")))
    args = parser.parse_args()
    build(args.input, args.output_dir)


if __name__ == "__main__":
    main()
