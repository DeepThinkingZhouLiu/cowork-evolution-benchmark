#!/usr/bin/env python3
"""受控购物清单题的专属 Oracle。"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree
from openpyxl import load_workbook


OUTPUT_NAME = "controlled_shopping_purchase_brief.docx"


@dataclass(frozen=True)
class Ticket:
    identifier: str
    filename: str
    item: str
    action: str
    before: Decimal
    after: Decimal
    source_location: str

    @property
    def delta(self) -> Decimal:
        return self.after - self.before


def number(value: Decimal) -> str:
    return str(int(value)) if value == value.to_integral() else f"{value.normalize()}"


def money(value: Decimal) -> str:
    return f"${value:.2f}"


def signed(value: Decimal) -> str:
    return f"+{number(value)}" if value >= 0 else number(value)


def read_baseline(path: Path) -> list[tuple[int, str, Decimal]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    assert [cell.value for cell in sheet[1]][:2] == ["Name", "amount"]
    rows = [(index, str(item), Decimal(str(amount))) for index, (item, amount) in enumerate(sheet.iter_rows(min_row=2, max_col=2, values_only=True), start=2)]
    workbook.close()
    assert rows == [(2, "apple", Decimal("2")), (3, "cabbage", Decimal("2")), (4, "onion", Decimal("2"))]
    return rows


def load_record(path: Path) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    assert isinstance(value.get("task"), str) and isinstance(value.get("evaluation"), list)
    return value


def truth(input_dir: Path):
    baseline = read_baseline(input_dir / "baseline_shopping_list.xlsx")
    baseline_map = OrderedDict((item, amount) for _, item, amount in baseline)

    t01 = load_record(input_dir / "change_T01_reduce_apple.json")
    assert "reduce half amount of apple" in t01["task"].lower()
    apple_target = Decimal(str(t01["evaluation"][0]["args"]["matches"][0]["value"]))
    assert apple_target == baseline_map["apple"] / 2

    t02 = load_record(input_dir / "change_T02_remove_onion.json")
    assert "remove this entry" in t02["task"].lower() and "onion" in t02["task"].lower()
    assert [str(value) for value in t02["evaluation"][0]["args"]["keywords"]] == ["onion", "2"]

    t03 = load_record(input_dir / "change_T03_add_yogurt.json")
    assert "add to shopping list" in t03["task"].lower()
    add_matches = t03["evaluation"][0]["args"]["matches"]
    added_item = str(next(match["value"] for match in add_matches if match["col"] == "1"))
    added_amount = Decimal(str(next(match["value"] for match in add_matches if match["col"] == "2")))

    pricing = load_record(input_dir / "pricing_policy.json")
    match = re.search(r"apple is\s*([0-9.]+)\$/pound,\s*others are\s*([0-9.]+)\$/pound", pricing["task"], re.I)
    assert match
    apple_rate, other_rate = map(Decimal, match.groups())

    tickets = [
        Ticket("T01", "change_T01_reduce_apple.json", "apple", "Reduce by half", baseline_map["apple"], apple_target, "evaluation[0].args.matches[0]"),
        Ticket("T02", "change_T02_remove_onion.json", "onion", "Remove — already at home", baseline_map["onion"], Decimal("0"), "task + evaluation[0].args.keywords"),
        Ticket("T03", "change_T03_add_yogurt.json", added_item, "Add approved item", Decimal("0"), added_amount, "evaluation[0].args.matches[0:2]"),
    ]

    current = OrderedDict(baseline_map)
    current["apple"] = apple_target
    del current["onion"]
    current[added_item] = added_amount
    rates = {item: apple_rate if item == "apple" else other_rate for item in current}
    costs = {item: amount * rates[item] for item, amount in current.items()}
    return baseline, tickets, current, rates, costs


def shade(cell, color: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    element = properties.find(qn("w:shd")) or OxmlElement("w:shd")
    element.set(qn("w:fill"), color)
    if element.getparent() is None:
        properties.append(element)


def no_split(row) -> None:
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def repeat_header(row) -> None:
    element = OxmlElement("w:tblHeader")
    element.set(qn("w:val"), "true")
    row._tr.get_or_add_trPr().append(element)


def borders(table, color: str = "A6B7C8") -> None:
    container = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        line = OxmlElement(f"w:{edge}")
        line.set(qn("w:val"), "single")
        line.set(qn("w:sz"), "4")
        line.set(qn("w:space"), "0")
        line.set(qn("w:color"), color)
        container.append(line)
    table._tbl.tblPr.append(container)


def set_cell(cell, value: str, bold: bool = False, size: float = 8.5) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(value)
    run.bold = bold
    run.font.size = Pt(size)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def header_row(table, labels: list[str]) -> None:
    repeat_header(table.rows[0])
    no_split(table.rows[0])
    for cell, label in zip(table.rows[0].cells, labels):
        set_cell(cell, label, bold=True, size=8)
        shade(cell, "385723")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)


def bookmark(paragraph, name: str, identifier: int) -> None:
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(identifier))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(identifier))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def internal_link(paragraph, label: str, anchor: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), anchor)
    run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    properties.extend([color, underline])
    run.append(properties)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)


def field(paragraph, instruction: str, display: str) -> None:
    node = OxmlElement("w:fldSimple")
    node.set(qn("w:instr"), instruction)
    run = OxmlElement("w:r")
    text = OxmlElement("w:t")
    text.text = display
    run.append(text)
    node.append(run)
    paragraph._p.append(node)


def heading(document: Document, label: str, level: int, anchor: str, identifier: int):
    paragraph = document.add_heading(label, level=level)
    bookmark(paragraph, anchor, identifier)
    return paragraph


def configure_styles(document: Document) -> None:
    normal = document.styles["Normal"]
    normal.font.name = "Liberation Sans"
    normal.font.size = Pt(9)
    normal.paragraph_format.space_after = Pt(3)
    title = document.styles["Title"]
    title.font.name = "Liberation Sans"
    title.font.size = Pt(22)
    title.font.bold = True
    title.font.color.rgb = RGBColor(56, 87, 35)
    for name, size, color in [("Heading 1", 14, "385723"), ("Heading 2", 11, "548235")]:
        style = document.styles[name]
        style.font.name = "Liberation Sans"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.keep_with_next = True


def configure_page(section, landscape: bool = False) -> None:
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.55)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        if section.page_width < section.page_height:
            section.page_width, section.page_height = section.page_height, section.page_width
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        if section.page_width > section.page_height:
            section.page_width, section.page_height = section.page_height, section.page_width


def running_parts(document: Document) -> None:
    for section in document.sections:
        section.header.is_linked_to_previous = False
        line = section.header.paragraphs[0]
        line.text = "CONTROLLED SHOPPING LIST  |  READY FOR HANDOFF"
        line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        line.runs[0].font.size = Pt(8)
        line.runs[0].font.bold = True
        line.runs[0].font.color.rgb = RGBColor(56, 87, 35)
        section.footer.is_linked_to_previous = False
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run("Controlled working paper  •  ")
        field(footer, " PAGE ", "1")
        footer.add_run(" / ")
        field(footer, " NUMPAGES ", "3")
        for run in footer.runs:
            run.font.size = Pt(8)


def inject_comments(path: Path, comments: list[tuple[str, str]]) -> None:
    wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    ns = {"w": wns}
    temporary = path.with_suffix(".comments.docx")
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    root = etree.fromstring(parts["word/document.xml"])
    comments_root = etree.Element(qn("w:comments"), nsmap={"w": wns})
    for comment_id, (anchor, note) in enumerate(comments):
        paragraphs = root.xpath("//w:p[contains(string(.), $anchor)]", namespaces=ns, anchor=anchor)
        assert len(paragraphs) == 1, (anchor, len(paragraphs))
        paragraph = paragraphs[0]
        runs = paragraph.xpath("./w:r[w:t]", namespaces=ns)
        assert runs
        start = OxmlElement("w:commentRangeStart")
        start.set(qn("w:id"), str(comment_id))
        end = OxmlElement("w:commentRangeEnd")
        end.set(qn("w:id"), str(comment_id))
        reference_run = OxmlElement("w:r")
        reference = OxmlElement("w:commentReference")
        reference.set(qn("w:id"), str(comment_id))
        reference_run.append(reference)
        first_index = paragraph.index(runs[0])
        last_index = paragraph.index(runs[-1])
        paragraph.insert(first_index, start)
        paragraph.insert(last_index + 2, end)
        paragraph.insert(last_index + 3, reference_run)
        comment = etree.SubElement(comments_root, qn("w:comment"))
        comment.set(qn("w:id"), str(comment_id))
        comment.set(qn("w:author"), "Purchase Control Review")
        comment_paragraph = etree.SubElement(comment, qn("w:p"))
        comment_run = etree.SubElement(comment_paragraph, qn("w:r"))
        etree.SubElement(comment_run, qn("w:t")).text = note
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    parts["word/comments.xml"] = etree.tostring(comments_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    content_types = etree.fromstring(parts["[Content_Types].xml"])
    override = etree.SubElement(content_types, f"{{{content_ns}}}Override")
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone="yes")
    relationships = etree.fromstring(parts["word/_rels/document.xml.rels"])
    used = {relationship.get("Id") for relationship in relationships}
    counter = 1
    while f"rId{counter}" in used:
        counter += 1
    relationship = etree.SubElement(relationships, f"{{{rel_ns}}}Relationship")
    relationship.set("Id", f"rId{counter}")
    relationship.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
    relationship.set("Target", "comments.xml")
    parts["word/_rels/document.xml.rels"] = etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone="yes")
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items():
            info = zipfile.ZipInfo(name, (2026, 9, 8, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = infos[name].external_attr if name in infos else 0o600 << 16
            target.writestr(info, payload)
    shutil.move(temporary, path)


def build(input_dir: Path, output_path: Path) -> None:
    baseline, tickets, current, rates, costs = truth(input_dir)
    baseline_total = sum(amount for _, _, amount in baseline)
    final_total = sum(current.values())
    total_cost = sum(costs.values())
    average = final_total / len(current)

    document = Document()
    configure_styles(document)
    configure_page(document.sections[0])
    properties = document.core_properties
    properties.title = "Controlled Shopping-List Purchase Brief"
    properties.author = "Purchase Control Review"
    properties.subject = "Baseline reconciliation, approved changes, and costed handoff"
    properties.keywords = "shopping list; change control; reconciliation; costing; handoff"

    title = document.add_paragraph(style="Title")
    title.add_run("Controlled Shopping-List\nPurchase Brief")
    subtitle = document.add_paragraph("Evidence-controlled household purchase handoff  |  Five frozen source files")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in subtitle.runs:
        run.italic = True
        run.font.color.rgb = RGBColor(89, 89, 89)

    status = document.add_table(rows=1, cols=2)
    set_cell(status.cell(0, 0), "CONTROL STATUS", bold=True, size=9)
    set_cell(status.cell(0, 1), "CONTROLLED — READY FOR HANDOFF", bold=True, size=10)
    shade(status.cell(0, 0), "385723")
    shade(status.cell(0, 1), "E2F0D9")
    status.cell(0, 0).paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    borders(status, "385723")

    heading(document, "Navigation", 1, "navigation", 1)
    navigation = document.add_paragraph()
    for index, (label, anchor) in enumerate([
        ("Executive reconciliation", "executive_reconciliation"),
        ("Current purchase list", "current_purchase"),
        ("Change ledger", "change_ledger"),
        ("Baseline reconciliation", "baseline_reconciliation"),
        ("Calculation notes", "calculation_notes"),
    ]):
        if index:
            navigation.add_run("  •  ")
        internal_link(navigation, label, anchor)
    toc = document.add_paragraph()
    toc.add_run("Native table of contents: ").bold = True
    field(toc, ' TOC \\o "1-2" \\h \\z \\u ', "Right-click and update field")

    heading(document, "Executive reconciliation", 1, "executive_reconciliation", 2)
    document.add_paragraph(
        "The baseline register was reconciled against three approved change records in ticket-ID order. One baseline quantity was reduced, one baseline item was removed but retained in the audit trail, one new item was appended, and one baseline item remained unchanged. Every current line can be costed under the frozen pricing policy, so the controlled list is ready for handoff."
    )
    facts = document.add_table(rows=3, cols=4)
    header_row(facts, ["Control metric", "Result", "Control metric", "Result"])
    metric_rows = [
        ("Baseline amount", number(baseline_total), "Final amount", number(final_total)),
        ("Signed amount change", signed(final_total - baseline_total), "Current line count", str(len(current))),
    ]
    for row, values in zip(facts.rows[1:], metric_rows):
        no_split(row)
        for cell, value in zip(row.cells, values):
            set_cell(cell, value, bold=value in {number(final_total), money(total_cost)})
    borders(facts)
    summary = document.add_paragraph()
    summary.add_run("Cost control: ").bold = True
    summary.add_run(f"estimated current-list total {money(total_cost)}; average final amount {average:.2f} lb per current line. ")
    internal_link(summary, "Open calculation notes", "calculation_notes")

    heading(document, "Source register", 1, "source_register", 3)
    source_table = document.add_table(rows=1, cols=3)
    header_row(source_table, ["Frozen source", "Control role", "Verification location"])
    sources = [
        ("baseline_shopping_list.xlsx", "Starting item/amount register", "Sheet1 rows 2–4"),
        ("change_T01_reduce_apple.json", "Approved quantity reduction", "task + evaluation[0].args.matches[0]"),
        ("change_T02_remove_onion.json", "Approved removal", "task + evaluation[0].args.keywords"),
        ("change_T03_add_yogurt.json", "Approved addition", "task + evaluation[0].args.matches[0:2]"),
        ("pricing_policy.json", "Named exception and other-item rate", "task + evaluation entries"),
    ]
    for filename, role, location in sources:
        row = source_table.add_row()
        for cell, value in zip(row.cells, [filename, role, location]):
            set_cell(cell, value, size=8)
        no_split(row)
    borders(source_table)

    landscape = document.add_section(WD_SECTION.NEW_PAGE)
    configure_page(landscape, landscape=True)
    heading(document, "Current purchase list", 1, "current_purchase", 10)
    document.add_paragraph("Only these lines are current. Surviving baseline lines keep their source order; the approved addition follows them. Amount is treated as pounds only for the policy-based estimate.")
    purchase = document.add_table(rows=1, cols=6)
    header_row(purchase, ["Item", "Final amount (lb)", "Unit rate", "Extended cost", "Change origin", "Evidence / control note"])
    origin = {"apple": "T01", "cabbage": "Baseline", "yogurt": "T03"}
    notes = {
        "apple": "T01 update: baseline amount halved; baseline Sheet1 row 2.",
        "cabbage": "Unchanged baseline line; baseline Sheet1 row 3.",
        "yogurt": "T03 addition: approved yogurt line; JSON evaluation matches.",
    }
    for item, amount in current.items():
        row = purchase.add_row()
        values = [item, number(amount), f"{money(rates[item])}/lb", money(costs[item]), origin[item], notes[item]]
        for cell, value in zip(row.cells, values):
            set_cell(cell, value, bold=item in {"apple", "yogurt"}, size=8)
        shade(row.cells[5], "FFF2CC" if item in {"apple", "yogurt"} else "FFFFFF")
        bookmark(row.cells[0].paragraphs[0], f"current_{item}", 20 + len(purchase.rows))
        no_split(row)
    total_row = purchase.add_row()
    for cell, value in zip(total_row.cells, ["TOTAL", number(final_total), "—", money(total_cost), "3 current lines", "Formula-checked"]):
        set_cell(cell, value, bold=True, size=8)
        shade(cell, "E2F0D9")
    no_split(total_row)
    borders(purchase)

    heading(document, "Approved change ledger", 1, "change_ledger", 30)
    ledger = document.add_table(rows=1, cols=8)
    header_row(ledger, ["Ticket", "Source file", "Item", "Action", "Before", "After", "Delta", "Disposition / source location"])
    for index, ticket in enumerate(tickets):
        row = ledger.add_row()
        values = [ticket.identifier, ticket.filename, ticket.item, ticket.action, number(ticket.before), number(ticket.after), signed(ticket.delta), f"Applied once • {ticket.source_location}"]
        for cell, value in zip(row.cells, values):
            set_cell(cell, value, size=7.7)
        bookmark(row.cells[0].paragraphs[0], f"ticket_{ticket.identifier.lower()}", 40 + index)
        no_split(row)
    borders(ledger)

    heading(document, "Baseline-to-final reconciliation", 1, "baseline_reconciliation", 50)
    reconciliation = document.add_table(rows=1, cols=7)
    header_row(reconciliation, ["Baseline row", "Item", "Baseline", "Final", "Delta", "State", "Controlling evidence"])
    baseline_lookup = {item: (row, amount) for row, item, amount in baseline}
    all_items = [item for _, item, _ in baseline] + [item for item in current if item not in baseline_lookup]
    state = {"apple": "updated", "cabbage": "unchanged", "onion": "removed", "yogurt": "added"}
    evidence = {"apple": "T01", "cabbage": "baseline only", "onion": "T02", "yogurt": "T03"}
    for item in all_items:
        row_number, before = baseline_lookup.get(item, (None, Decimal("0")))
        after = current.get(item, Decimal("0"))
        row = reconciliation.add_row()
        values = [str(row_number) if row_number else "—", item, number(before), number(after), signed(after - before), state[item], evidence[item]]
        for cell, value in zip(row.cells, values):
            set_cell(cell, value, size=8)
        shade(row.cells[5], {"updated": "FFF2CC", "unchanged": "E2F0D9", "removed": "F4CCCC", "added": "DDEBF7"}[state[item]])
        no_split(row)
    borders(reconciliation)

    portrait = document.add_section(WD_SECTION.NEW_PAGE)
    configure_page(portrait)
    heading(document, "Calculation and handoff notes", 1, "calculation_notes", 60)
    calculations = document.add_table(rows=1, cols=3)
    header_row(calculations, ["Control", "Calculation", "Result"])
    rows = [
        ("Baseline total amount", "2 + 2 + 2", number(baseline_total)),
        ("Final total amount", "1 + 2 + 2", number(final_total)),
        ("Signed amount change", "final 5 − baseline 6", signed(final_total - baseline_total)),
        ("Average final amount", "final amount 5 ÷ 3 current lines", f"{average:.2f} lb"),
        ("Estimated cost", "$2.00 + $3.00 + $3.00", money(total_cost)),
    ]
    for label, calculation, result in rows:
        row = calculations.add_row()
        for cell, value in zip(row.cells, [label, calculation, result]):
            set_cell(cell, value)
        no_split(row)
    borders(calculations)

    heading(document, "Method and evidence boundary", 1, "evidence_boundary", 70)
    for text in [
        "Application method: start with the baseline, apply T01, T02, and T03 once each in ID order, retain surviving baseline order, then append the approved new item.",
        "Pricing method: apple uses the named $2.00-per-pound exception; every other final item uses the policy's $1.50-per-pound rate. Amount is interpreted as pounds only because the pricing rule is expressed per pound.",
        "The dollar figures are policy-based estimates, not a store quotation or price guarantee. No store, vendor, brand, tax, discount, budget, package size, or purchase date is supplied, so none is asserted here.",
        "Handoff control: buy only the current-list lines. Onion remains visible solely in the change and reconciliation ledgers as a removed baseline item.",
    ]:
        document.add_paragraph(text, style="List Bullet")
    links = document.add_paragraph("Evidence shortcuts: ")
    internal_link(links, "apple / T01", "ticket_t01")
    links.add_run("  •  ")
    internal_link(links, "onion / T02", "ticket_t02")
    links.add_run("  •  ")
    internal_link(links, "yogurt / T03", "ticket_t03")
    links.add_run("  •  ")
    internal_link(links, "current list", "current_purchase")

    running_parts(document)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)
    inject_comments(
        output_path,
        [
            ("T01 update: baseline amount halved", "T01 control evidence: change_T01_reduce_apple.json, task and evaluation[0].args.matches[0], applied to baseline_shopping_list.xlsx Sheet1 row 2."),
            ("T03 addition: approved yogurt line", "T03 control evidence: change_T03_add_yogurt.json, task and evaluation[0].args.matches[0:2], appended after surviving baseline lines."),
        ],
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    arguments = parser.parse_args()
    build(arguments.input_dir, arguments.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
