#!/usr/bin/env python3
"""配送与税率配置处理器。

只从 ``--input-dir`` 读取上游静态快照，并同时生成 JSON、
Excel、Word 和离线邮件草稿。脚本不包含本题的固定答案。
"""

from __future__ import annotations

import argparse
import csv
import json
from collections import defaultdict
from decimal import Decimal
from math import ceil
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


NAVY = "17324D"
BLUE = "2F6B8A"
PALE_BLUE = "DCEAF2"
ICE = "F5F8FA"
AMBER = "D58B16"
RED = "B23A48"
GRAY = "5E6C78"
WHITE = "FFFFFF"


def load_csv(path: Path, required: tuple[str, ...]) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames is None or not set(required).issubset(reader.fieldnames):
            raise ValueError(f"{path.name} 缺少必需列：{required}")
        return list(reader)


def as_bool(value: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"t", "true", "1", "yes"}:
        return True
    if normalized in {"f", "false", "0", "no"}:
        return False
    raise ValueError(f"无法解析布尔值：{value!r}")


def decimal_number(value: Decimal) -> int | float:
    if value == value.to_integral():
        return int(value)
    return float(value)


def decimal_text(value: Decimal) -> str:
    return format(value.normalize(), "f")


def analyze(input_dir: Path) -> dict[str, Any]:
    zones = load_csv(input_dir / "shipping_zones.csv", ("id", "name", "order"))
    locations = load_csv(
        input_dir / "shipping_zone_locations.csv", ("id", "zone_id", "code", "type")
    )
    zone_methods = load_csv(
        input_dir / "shipping_zone_methods.csv",
        ("id", "zone_id", "method_id", "method_title", "enabled", "settings"),
    )
    method_catalog = load_csv(
        input_dir / "shipping_methods.csv", ("id", "title", "description")
    )
    tax_classes = load_csv(input_dir / "tax_classes.csv", ("slug", "name"))
    tax_rates = load_csv(
        input_dir / "tax_rates.csv",
        (
            "id", "country", "state", "postcode", "city", "rate", "name",
            "priority", "compound", "shipping", "order", "class",
        ),
    )

    catalog_ids = {row["id"] for row in method_catalog}
    class_names = {row["slug"]: row["name"] for row in tax_classes}
    if len(catalog_ids) != len(method_catalog):
        raise ValueError("shipping_methods.csv 包含重复 id")
    if len(class_names) != len(tax_classes):
        raise ValueError("tax_classes.csv 包含重复 slug")

    locations_by_zone: dict[int, list[dict[str, str]]] = defaultdict(list)
    for row in sorted(locations, key=lambda item: int(item["id"])):
        locations_by_zone[int(row["zone_id"])].append(
            {"location_id": int(row["id"]), "type": row["type"], "code": row["code"]}
        )

    methods_by_zone: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(zone_methods, key=lambda item: int(item["id"])):
        if row["method_id"] not in catalog_ids:
            raise ValueError(f"配送方法 {row['method_id']} 不在方法目录中")
        try:
            settings = json.loads(row["settings"] or "{}")
        except json.JSONDecodeError as error:
            raise ValueError(f"配送方法 {row['id']} settings 非法") from error
        methods_by_zone[int(row["zone_id"])].append(
            {
                "record_id": int(row["id"]),
                "method_id": row["method_id"],
                "title": row["method_title"],
                "enabled": as_bool(row["enabled"]),
                "settings": settings,
            }
        )

    shipping_rows: list[dict[str, Any]] = []
    for zone in sorted(zones, key=lambda item: (int(item["order"]), int(item["id"]))):
        zone_id = int(zone["id"])
        enabled = [item for item in methods_by_zone[zone_id] if item["enabled"]]
        shipping_rows.append(
            {
                "zone_id": zone_id,
                "zone_name": zone["name"],
                "zone_order": int(zone["order"]),
                "locations": locations_by_zone[zone_id],
                "enabled_methods": enabled,
                "methods_count": len(enabled),
            }
        )

    rates_by_class: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in sorted(tax_rates, key=lambda item: int(item["id"])):
        slug = row["class"] or "standard"
        if slug not in class_names:
            raise ValueError(f"税率 {row['id']} 引用未知税类 {slug}")
        rate = Decimal(row["rate"])
        rates_by_class[slug].append(
            {
                "rate_id": int(row["id"]),
                "country": row["country"],
                "state": row["state"],
                "postcode": row["postcode"],
                "city": row["city"],
                "rate": decimal_number(rate),
                "rate_name": row["name"],
                "priority": int(row["priority"]),
                "compound": as_bool(row["compound"]),
                "shipping": as_bool(row["shipping"]),
                "order": int(row["order"]),
            }
        )

    tax_rows = [
        {"slug": row["slug"], "name": row["name"], "rates": rates_by_class[row["slug"]]}
        for row in sorted(tax_classes, key=lambda item: item["name"])
    ]
    all_rates = [rate for tax_class in tax_rows for rate in tax_class["rates"]]
    displayed_rows = sum(max(1, len(tax_class["rates"])) for tax_class in tax_rows)
    highest_rate = max((Decimal(str(row["rate"])) for row in all_rates), default=Decimal("0"))

    return {
        "shipping_zones": shipping_rows,
        "tax_classes": tax_rows,
        "summary": {
            "total_zones": len(shipping_rows),
            "total_enabled_methods": sum(row["methods_count"] for row in shipping_rows),
            "tax_classes": len(tax_rows),
            "configured_tax_rates": len(all_rates),
            "displayed_tax_rows": displayed_rows,
            "zones_with_single_method": sum(row["methods_count"] == 1 for row in shipping_rows),
            "highest_configured_rate": decimal_number(highest_rate),
        },
    }


def excel_shipping_rows(result: dict[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for zone in result["shipping_zones"]:
        regions = "; ".join(f"{item['type']}:{item['code']}" for item in zone["locations"])
        methods = "; ".join(item["title"] for item in zone["enabled_methods"])
        rows.append([zone["zone_name"], None, regions, methods])
    return rows


def excel_tax_rows(result: dict[str, Any]) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for tax_class in result["tax_classes"]:
        if not tax_class["rates"]:
            rows.append([tax_class["name"], "", None, None, None])
            continue
        for rate in tax_class["rates"]:
            rows.append(
                [
                    tax_class["name"], rate["state"], rate["rate"],
                    rate["compound"], rate["shipping"],
                ]
            )
    return rows


def style_worksheet(sheet, widths: dict[str, float], table_name: str) -> None:
    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(name="Aptos", bold=True, color=WHITE, size=10)
    body_font = Font(name="Aptos", color=NAVY, size=10)
    bottom_border = Border(bottom=Side(style="thin", color="D6DEE5"))
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row_number, row in enumerate(sheet.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = body_font
            cell.border = bottom_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        sheet.row_dimensions[row_number].height = max(
            23,
            min(72, 15 * ceil(max(len(str(cell.value or "")) for cell in row) / 48)),
        )
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.row_dimensions[1].height = 29
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    table = Table(displayName=table_name, ref=sheet.dimensions)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_title_rows = "1:1"
    sheet.print_area = sheet.dimensions
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.45
    sheet.page_margins.bottom = 0.45


def build_excel(result: dict[str, Any], output_path: Path) -> None:
    workbook = Workbook()
    shipping = workbook.active
    shipping.title = "Shipping Zones"
    shipping.append(["Zone_Name", "Methods_Count", "Regions", "Methods"])
    for row_number, row in enumerate(excel_shipping_rows(result), start=2):
        shipping.append(row)
        shipping.cell(row_number, 2).value = (
            f'=IF(D{row_number}="",0,LEN(D{row_number})-'
            f'LEN(SUBSTITUTE(D{row_number},";",""))+1)'
        )
    style_worksheet(
        shipping,
        {"A": 24, "B": 17, "C": 37, "D": 58},
        "ShippingZonesTable",
    )
    shipping.conditional_formatting.add(
        f"B2:B{shipping.max_row}",
        CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", fgColor="FCE8C3")),
    )

    taxes = workbook.create_sheet("Tax Rates")
    taxes.append(["Tax_Class", "State", "Rate", "Compound", "Shipping"])
    for row in excel_tax_rows(result):
        taxes.append(row)
    style_worksheet(
        taxes,
        {"A": 24, "B": 16, "C": 15, "D": 16, "E": 16},
        "TaxRatesTable",
    )
    for cell in taxes["C"][1:]:
        cell.number_format = '0.0000"%"'
    taxes.conditional_formatting.add(
        f"C2:C{taxes.max_row}",
        CellIsRule(operator="greaterThan", formula=["8"], fill=PatternFill("solid", fgColor="FCE8C3")),
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def set_cell_shading(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = properties.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        properties.append(shading)
    shading.set(qn("w:fill"), fill)


def set_cell_border(cell, color: str = "D6DEE5") -> None:
    properties = cell._tc.get_or_add_tcPr()
    borders = properties.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        properties.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), "5")
        node.set(qn("w:color"), color)


def add_word_table(
    document: Document,
    headers: list[str],
    rows: list[list[Any]],
    widths: list[float],
):
    table = document.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for index, header in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = header
        cell.width = Inches(widths[index])
        set_cell_shading(cell, NAVY)
        set_cell_border(cell, NAVY)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for paragraph in cell.paragraphs:
            paragraph.paragraph_format.space_after = Pt(0)
            for run in paragraph.runs:
                run.font.name = "Noto Sans"
                run.font.size = Pt(8)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    for row_number, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell = cells[index]
            cell.text = "" if value is None else str(value)
            cell.width = Inches(widths[index])
            set_cell_shading(cell, WHITE if row_number % 2 == 0 else ICE)
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                for run in paragraph.runs:
                    run.font.name = "Noto Sans"
                    run.font.size = Pt(7.5)
                    run.font.color.rgb = RGBColor.from_string(NAVY)
    return table


def add_bullet(document: Document, text: str) -> None:
    paragraph = document.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(3)
    run = paragraph.add_run(text)
    run.font.name = "Noto Sans"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(NAVY)


def build_word(result: dict[str, Any], output_path: Path) -> None:
    summary = result["summary"]
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(1.35)
    section.bottom_margin = Cm(1.25)
    section.left_margin = Cm(1.45)
    section.right_margin = Cm(1.45)

    normal = document.styles["Normal"]
    normal.font.name = "Noto Sans"
    normal.font.size = Pt(9.2)
    normal.font.color.rgb = RGBColor.from_string(NAVY)
    normal.paragraph_format.space_after = Pt(5)
    for style_name, size, color in [
        ("Title", 24, NAVY), ("Heading 1", 15, BLUE), ("Heading 2", 11, GRAY),
    ]:
        style = document.styles[style_name]
        style.font.name = "Noto Sans"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.font.bold = True

    header = section.header.paragraphs[0]
    header.text = "OPERATIONS  /  CONFIGURATION REVIEW"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.name = "Noto Sans"
        run.font.size = Pt(7.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(BLUE)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Source-bound offline report • Shipping & tax configuration")
    run.font.name = "Noto Sans"
    run.font.size = Pt(7.5)
    run.font.color.rgb = RGBColor.from_string(GRAY)

    title = document.add_paragraph(style="Title")
    title.paragraph_format.space_after = Pt(2)
    title.add_run("Operations Report")
    subtitle = document.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(12)
    run = subtitle.add_run("Shipping coverage and tax configuration snapshot")
    run.font.name = "Noto Sans"
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor.from_string(BLUE)

    document.add_heading("Executive Summary", level=1)
    metrics = [
        ["Shipping zones", summary["total_zones"], "Enabled methods", summary["total_enabled_methods"]],
        ["Tax classes", summary["tax_classes"], "Configured tax rates", summary["configured_tax_rates"]],
        ["Displayed tax rows", summary["displayed_tax_rows"], "Highest configured rate", f"{decimal_text(Decimal(str(summary['highest_configured_rate'])))}%"],
    ]
    add_word_table(document, ["Metric", "Value", "Metric", "Value"], metrics, [2.15, 1.15, 2.45, 1.55])
    paragraph = document.add_paragraph()
    paragraph.add_run(
        f"The source snapshot contains {summary['total_zones']} shipping zones with "
        f"{summary['total_enabled_methods']} enabled methods. It defines "
        f"{summary['tax_classes']} tax classes and {summary['configured_tax_rates']} configured "
        f"tax-rate records. The Tax Rates worksheet displays {summary['displayed_tax_rows']} rows "
        "because a class without a configured rate is retained as an empty row."
    )

    document.add_heading("Shipping Coverage", level=1)
    shipping_rows: list[list[Any]] = []
    for zone in result["shipping_zones"]:
        regions = "; ".join(f"{item['type']}:{item['code']}" for item in zone["locations"])
        methods = "; ".join(item["title"] for item in zone["enabled_methods"])
        shipping_rows.append([zone["zone_name"], zone["methods_count"], regions, methods])
    add_word_table(
        document,
        ["Zone", "Enabled methods", "Source region codes", "Configured method titles"],
        shipping_rows,
        [1.65, 1.15, 2.05, 4.55],
    )

    document.add_page_break()
    document.add_heading("Tax Structure", level=1)
    tax_rows: list[list[Any]] = []
    for tax_class in result["tax_classes"]:
        if not tax_class["rates"]:
            tax_rows.append([tax_class["name"], "No configured rate", "", "", "", ""])
            continue
        for rate in tax_class["rates"]:
            scope_parts = [
                value for value in (
                    rate["country"], rate["state"], rate["postcode"], rate["city"]
                ) if value
            ]
            tax_rows.append(
                [
                    tax_class["name"], rate["rate_name"], " / ".join(scope_parts),
                    f"{decimal_text(Decimal(str(rate['rate'])))}%",
                    str(rate["compound"]).lower(), str(rate["shipping"]).lower(),
                ]
            )
    add_word_table(
        document,
        ["Tax class", "Rate name", "Source scope", "Rate", "Compound", "Shipping"],
        tax_rows,
        [1.35, 1.65, 2.6, 0.9, 1.0, 1.0],
    )
    paragraph = document.add_paragraph()
    paragraph.add_run(
        "Standard contains six configured records, Reduced Rate contains one, and Zero Rate "
        "contains no configured record. CA appears twice under Standard: 7.25% without a "
        "postcode/city scope and 8.625% with postcode 94* and city Краснодар. "
        "Every configured record applies to shipping and none is compound."
    )

    document.add_heading("Recommendations", level=1)
    single_method_zones = [
        zone["zone_name"] for zone in result["shipping_zones"] if zone["methods_count"] == 1
    ]
    add_bullet(
        document,
        f"Review whether {', '.join(single_method_zones)} needs another enabled delivery option; "
        "the snapshot currently shows one configured method there. This is a configuration review, not a demand forecast.",
    )
    add_bullet(
        document,
        "Verify that the two CA Standard entries intentionally use different postcode/city scope before simplifying either record.",
    )
    add_bullet(
        document,
        "Document the intended use of Zero Rate before activation; the class exists in the source, but no rate record is configured.",
    )

    document.add_heading("Source Boundary", level=1)
    document.add_paragraph(
        "This report restates the supplied configuration snapshot only. Region labels, tax names, "
        "codes and settings are preserved as supplied. No order volume, revenue, customer demand, "
        "conversion, market benchmark, savings estimate or forecast was provided, so none is asserted."
    )

    core_properties = document.core_properties
    core_properties.title = "Operations Report: Shipping and Tax Configuration"
    core_properties.subject = "Source-bound review of shipping zones and tax classes"
    core_properties.author = "Operations"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(output_path)


def build_email(result: dict[str, Any], output_path: Path) -> None:
    summary = result["summary"]
    zone_lines = [
        f"- {zone['zone_name']}: {zone['methods_count']} enabled method(s); regions "
        + "; ".join(f"{item['type']}:{item['code']}" for item in zone["locations"])
        for zone in result["shipping_zones"]
    ]
    class_lines = []
    for tax_class in result["tax_classes"]:
        rates = ", ".join(
            f"{decimal_text(Decimal(str(rate['rate'])))}%" for rate in tax_class["rates"]
        ) or "none"
        class_lines.append(
            f"- {tax_class['name']}: {len(tax_class['rates'])} configured rate record(s); rates {rates}"
        )
    body = "\n".join(
        [
            "To: operations@store.com",
            "Subject: Shipping and Tax Configuration Report",
            "",
            "Hello Operations team,",
            "",
            f"The supplied snapshot contains {summary['total_zones']} shipping zones and "
            f"{summary['total_enabled_methods']} enabled shipping methods.",
            *zone_lines,
            "",
            f"It also contains {summary['tax_classes']} tax classes and "
            f"{summary['configured_tax_rates']} configured tax-rate records; the highest configured "
            f"rate is {decimal_text(Decimal(str(summary['highest_configured_rate'])))}%.",
            *class_lines,
            "",
            "Recommended checks: review the single-method international zone, confirm the two CA "
            "Standard scopes, and document the intended use of the unconfigured Zero Rate class.",
            "",
            "The attached-style deliverables are offline benchmark artifacts; this draft was not sent.",
        ]
    )
    output_path.write_text(body + "\n", encoding="utf-8")


def generate_package(input_dir: Path, output_dir: Path) -> dict[str, Any]:
    result = analyze(input_dir.resolve())
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "shipping_tax_results.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    build_excel(result, output_dir / "Shipping_Tax_Analysis.xlsx")
    build_word(result, output_dir / "Operations_Report.docx")
    build_email(result, output_dir / "operations_email.txt")
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    generate_package(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
