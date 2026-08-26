#!/usr/bin/env python3
"""officeval_036 的逐题参考解法。

48 行 PDF 记录是逐行阅读原始两页银行对账单后转录的；
输出保留 Word 模板中已有的 2 行真实样例，共 50 行。
这不是可替换输入来批量造题的 PDF 抽取器。
"""

from __future__ import annotations

import argparse
from decimal import Decimal
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


INPUT_NAME = "public_ready_statement.docx"
OUTPUT_NAME = "public_ready_statement_收支记录整理.docx"
HEADER = ["Date", "Description", "Withdrawals", "Deposits", "Balance"]
CELL_WIDTH_TWIPS = 2086  # 3.68 cm
DATA_ROW_HEIGHT_CM = 1.80

# Word 模板中已有的两条记录，保留事实，只统一金额格式。
TEMPLATE_RECORDS = [
    ("Mar 22", "Direct deposit from Cedarline Retail Group", "", "$8,732.40", "$9,016.55"),
    ("Mar 25", "Vendor settlement - 出支记录", "$6,918.75", "", "$2,097.80"),
]

# 按 PDF 两页中的日期和出现顺序逐行转录：1 行期初余额、46 笔收支、1 行期末汇总。
PDF_RECORDS = [
    ("01 Dec.", "Balance forward", "", "", "$43,712.48"),
    ("01 Dec.", "Insurance premium - studio toolkits", "$218.40", "", "$43,494.08"),
    ("01 Dec.", "Supplier payout - compact label printers", "$184.75", "", "$43,309.33"),
    ("02 Dec.", "Merchant batch - bamboo organizer trays", "", "$6,870.20", "$50,179.53"),
    ("02 Dec.", "Merchant batch - LED task lamps", "", "$7,940.00", "$58,119.53"),
    ("02 Dec.", "Merchant batch - ceramic storage jars", "", "$2,215.35", "$60,334.88"),
    ("03 Dec.", "Client transfer - folding display stands", "", "$1,135.70", "$61,470.58"),
    ("05 Dec.", "Marketplace settlement - wool desk mats", "", "$695.42", "$62,166.00"),
    ("09 Dec.", "Platform fee - travel cable pouches", "$63.25", "", "$62,102.75"),
    ("10 Dec.", "Client transfer - glass water carafes", "", "$3,420.00", "$65,522.75"),
    ("11 Dec.", "Utilities payment - workshop power", "$562.90", "", "$64,959.85"),
    ("12 Dec.", "Payment gateway batch - recycled notebooks", "", "$11,780.50", "$76,740.35"),
    ("15 Dec.", "Client transfer - countertop spice racks", "", "$2,640.00", "$79,380.35"),
    ("15 Dec.", "Cheque cleared - portable shelving units", "$3,496.80", "", "$75,883.55"),
    ("15 Dec.", "ATM cash deposit - weekend sales", "", "$7,530.00", "$83,413.55"),
    ("15 Dec.", "Remote cheque deposit - acoustic wall panels", "", "$615.90", "$84,029.45"),
    ("15 Dec.", "Cheque cleared - magnetic whiteboards", "$2,745.60", "", "$81,283.85"),
    ("15 Dec.", "Bill payment - warehouse heat", "$684.30", "", "$80,599.55"),
    ("15 Dec.", "Cheque cleared - washable floor runners", "$2,430.00", "", "$78,169.55"),
    ("15 Dec.", "Cheque cleared - adjustable coat racks", "$3,215.45", "", "$74,954.10"),
    ("15 Dec.", "Cheque cleared - kraft mailer cartons", "$1,955.20", "", "$72,998.90"),
    ("15 Dec.", "Cheque cleared - countertop scales", "$3,118.70", "", "$69,880.20"),
    ("17 Dec.", "Supplier transfer - insulated lunch tins", "$1,022.50", "", "$68,857.70"),
    ("17 Dec.", "Supplier transfer - silicone cable clips", "$742.80", "", "$68,114.90"),
    ("17 Dec.", "Supplier transfer - walnut monitor risers", "$1,604.40", "", "$66,510.50"),
    ("18 Dec.", "Supplier transfer - linen storage bins", "$1,885.00", "", "$64,625.50"),
    ("18 Dec.", "Cheque cleared - reusable shipping crates", "$2,875.35", "", "$61,750.15"),
    ("19 Dec.", "Remote cheque deposit - desk organizers", "", "$2,940.75", "$64,690.90"),
    ("23 Dec.", "Payment gateway batch - clip-on reading lights", "", "$8,210.60", "$72,901.50"),
    ("24 Dec.", "Wholesale payment - cotton tote sets", "", "$5,476.25", "$78,377.75"),
    ("24 Dec.", "ATM deposit - cash sales", "", "$6,800.00", "$85,177.75"),
    ("29 Dec.", "Client transfer - enamel plant tags", "", "$2,125.00", "$87,302.75"),
    ("29 Dec.", "Client transfer - cork message boards", "", "$3,480.00", "$90,782.75"),
    ("29 Dec.", "Marketplace settlement - soft-close storage boxes", "", "$2,870.65", "$93,653.40"),
    ("29 Dec.", "Marketplace settlement - acrylic recipe holders", "", "$1,120.30", "$94,773.70"),
    ("30 Dec.", "Payment gateway batch - glass spice mills", "", "$4,435.40", "$99,209.10"),
    ("30 Dec.", "Cheque cleared - collapsible laundry carts", "$9,750.00", "", "$89,459.10"),
    ("30 Dec.", "Platform fee - leather key trays", "$483.60", "", "$88,975.50"),
    ("30 Dec.", "Platform fee - brushed metal hooks", "$37.25", "", "$88,938.25"),
    ("30 Dec.", "Platform fee - canvas tool rolls", "$52.18", "", "$88,886.07"),
    ("31 Dec.", "Cheque cleared - cedar drawer inserts", "$3,120.70", "", "$85,765.37"),
    ("31 Dec.", "Cheque cleared - pegboard panels", "$2,948.30", "", "$82,817.07"),
    ("31 Dec.", "Cheque cleared - wire baskets", "$2,104.90", "", "$80,712.17"),
    ("31 Dec.", "Cheque cleared - rolling trolleys", "$3,302.75", "", "$77,409.42"),
    ("31 Dec.", "Cheque cleared - cable trays", "$3,005.45", "", "$74,403.97"),
    ("31 Dec.", "Cheque cleared - bamboo file holders", "$3,260.20", "", "$71,143.77"),
    ("31 Dec.", "Monthly service fee - operating account", "$21.00", "", "$71,122.77"),
    ("31 Dec.", "Closing totals", "$54,891.73", "$82,302.02", "$71,122.77"),
]


def money_value(text: str) -> Decimal:
    return Decimal(text.replace("$", "").replace(",", ""))


def validate_transcription() -> None:
    if len(PDF_RECORDS) != 48:
        raise ValueError(f"PDF 转录应为 48 行，实际 {len(PDF_RECORDS)} 行")
    withdrawals = sum(
        (money_value(row[2]) for row in PDF_RECORDS[1:-1] if row[2]),
        Decimal("0"),
    )
    deposits = sum(
        (money_value(row[3]) for row in PDF_RECORDS[1:-1] if row[3]),
        Decimal("0"),
    )
    closing = PDF_RECORDS[-1]
    if withdrawals != money_value(closing[2]) or deposits != money_value(closing[3]):
        raise ValueError(
            f"转录合计与 PDF 期末汇总不符：{withdrawals=}, {deposits=}"
        )


def set_cell_width(cell, width_twips: int = CELL_WIDTH_TWIPS) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.first_child_found_in("w:tcW")
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width_twips))
    tc_w.set(qn("w:type"), "dxa")


def set_bottom_border(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    bottom = borders.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        borders.append(bottom)
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")  # 1/8 pt 为单位，6 = 0.75 pt
    bottom.set(qn("w:space"), "0")
    bottom.set(qn("w:color"), "D9D9D9")


def set_row_flags(row, *, repeat_header: bool = False) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    for tag in ("w:tblHeader", "w:cantSplit"):
        existing = tr_pr.find(qn(tag))
        if existing is not None:
            tr_pr.remove(existing)
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)
    if repeat_header:
        table_header = OxmlElement("w:tblHeader")
        table_header.set(qn("w:val"), "true")
        tr_pr.append(table_header)


def set_run_font(run, size_pt: float, *, bold: bool) -> None:
    run.font.name = "Arial"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)
    r_fonts = run._element.get_or_add_rPr().get_or_add_rFonts()
    for attribute in ("w:ascii", "w:hAnsi", "w:eastAsia"):
        r_fonts.set(qn(attribute), "Arial")


def write_cell(cell, text: str, *, header: bool, column_index: int) -> None:
    cell.text = ""
    cell.width = Cm(3.68)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_width(cell)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if column_index < 2 else WD_ALIGN_PARAGRAPH.RIGHT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0 if header else 10)
    paragraph.paragraph_format.line_spacing = 1.0 if header else 1.15
    run = paragraph.add_run(text)
    set_run_font(run, 11.5 if header else 10.5, bold=header or (not header and column_index == 4))
    if not header:
        set_bottom_border(cell)


def rebuild_record_table(document: Document) -> None:
    if len(document.tables) < 2:
        raise ValueError("模板中缺少收支记录表")
    table = document.tables[1]
    if [cell.text.strip() for cell in table.rows[0].cells] != HEADER:
        raise ValueError("第一个收支记录表表头与预期不符")
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[-1]._tr)

    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for grid_col in table._tbl.tblGrid.gridCol_lst:
        grid_col.set(qn("w:w"), str(CELL_WIDTH_TWIPS))

    header_row = table.rows[0]
    header_row.height = Cm(1.0)
    header_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    set_row_flags(header_row, repeat_header=True)
    for column_index, (cell, text) in enumerate(zip(header_row.cells, HEADER)):
        write_cell(cell, text, header=True, column_index=column_index)

    for record in [*TEMPLATE_RECORDS, *PDF_RECORDS]:
        row = table.add_row()
        row.height = Cm(DATA_ROW_HEIGHT_CM)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        set_row_flags(row)
        for column_index, (cell, text) in enumerate(zip(row.cells, record)):
            write_cell(cell, text, header=False, column_index=column_index)


def solve(input_path: Path, output_path: Path) -> None:
    validate_transcription()
    document = Document(input_path)
    rebuild_record_table(document)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir / INPUT_NAME, args.output_dir / OUTPUT_NAME)


if __name__ == "__main__":
    main()
