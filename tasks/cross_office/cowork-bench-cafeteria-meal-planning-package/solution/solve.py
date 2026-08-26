#!/usr/bin/env python3
"""从可回放 JSON 生成 Excel、Word 和 ICS Golden。"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


PROCESSOR_NAME = "meal_cost_optimizer.py"
JSON_NAME = "meal_optimization_results.json"
XLSX_NAME = "Weekly_Meal_Budget.xlsx"
DOCX_NAME = "Cafeteria_Meal_Proposal.docx"
ICS_NAME = "meal_prep_calendar.ics"
NAVY = "17365D"
BLUE = "2F75B5"
CYAN = "22A7B8"
PALE = "EAF2F8"
WHITE = "FFFFFF"
INK = "1F2937"
MUTED = "5B6B7F"


def set_cell_shading(cell, color: str) -> None:
    props = cell._tc.get_or_add_tcPr()
    shading = props.find(qn("w:shd"))
    if shading is None:
        shading = OxmlElement("w:shd")
        props.append(shading)
    shading.set(qn("w:fill"), color)


def set_cell_text(cell, text: str, bold: bool = False, color: str = INK, size: int = 9) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_excel_table(sheet, ref: str, name: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)


def style_sheet(sheet, widths: dict[str, float]) -> None:
    header_fill = PatternFill("solid", fgColor=NAVY)
    thin = Side(style="thin", color="D7E0EA")
    for cell in sheet[1]:
        cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 34
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Arial", size=9, color=INK)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)
        sheet.row_dimensions[row[0].row].height = 24
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A3
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.page_margins.left = 0.25
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.35
    sheet.page_margins.bottom = 0.35
    sheet.oddFooter.center.text = "Cowork-Bench · Cafeteria Meal Planning"


def write_workbook(path: Path, result: dict) -> None:
    workbook = Workbook()
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"

    prices = workbook.active
    prices.title = "Grocery_Prices"
    price_headers = ["Item", "Unit", "FreshMart_Price", "ValueGrocer_Price", "BulkBarn_Price", "Cheapest_Store", "Cheapest_Price"]
    prices.append(price_headers)
    for index, item in enumerate(result["price_rows"], start=2):
        prices.append([item[header] for header in price_headers[:5]] + [None, None])
        prices[f"F{index}"] = f'=SUBSTITUTE(INDEX($C$1:$E$1,1,MATCH(MIN(C{index}:E{index}),C{index}:E{index},0)),"_Price","")'
        prices[f"G{index}"] = f"=MIN(C{index}:E{index})"
    style_sheet(prices, {"A": 27, "B": 14, "C": 18, "D": 20, "E": 18, "F": 19, "G": 17})
    for row in prices.iter_rows(min_row=2, min_col=3, max_col=5):
        for cell in row:
            cell.number_format = '#,##0.00" ₽"'
    for cell in prices["G"][1:]:
        cell.number_format = '#,##0.00" ₽"'
    prices.conditional_formatting.add(f"G2:G{prices.max_row}", ColorScaleRule(start_type="min", start_color="63BE7B", end_type="max", end_color="F8696B"))
    add_excel_table(prices, f"A1:G{prices.max_row}", "Meal_Grocery_Prices")

    recipes = workbook.create_sheet("Recipe_Cost_Analysis")
    recipe_headers = ["Recipe_ID", "Recipe", "Matched_Ingredient_Count", "Unpriced_Ingredient_Count", "Total_Cost", "Cheapest_Store_Mix", "Cost_Per_Serving", "Unpriced_Ingredients"]
    recipes.append(recipe_headers)
    for index, item in enumerate(result["selected_recipes"], start=2):
        recipes.append(
            [
                item["Recipe_ID"], item["Recipe"], item["Matched_Ingredient_Count"], item["Unpriced_Ingredient_Count"],
                item["Total_Cost"], "/".join(item["Cheapest_Store_Mix"]), None, "; ".join(item["Unpriced_Ingredients"]),
            ]
        )
        recipes[f"G{index}"] = f"=ROUND(E{index}/4,2)"
    style_sheet(recipes, {"A": 34, "B": 28, "C": 18, "D": 19, "E": 15, "F": 31, "G": 18, "H": 43})
    for column in ("E", "G"):
        for cell in recipes[column][1:]:
            cell.number_format = '#,##0.00" ₽"'
    add_excel_table(recipes, f"A1:H{recipes.max_row}", "Meal_Recipe_Costs")

    plan = workbook.create_sheet("Weekly_Meal_Plan")
    plan_headers = ["Date", "Day", "Meal", "Recipe_ID", "Recipe", "Estimated_Cost"]
    plan.append(plan_headers)
    for index, item in enumerate(result["weekly_meal_plan"], start=2):
        plan.append([item[header] for header in plan_headers[:5]] + [None])
        plan[f"F{index}"] = f'=VLOOKUP(D{index},Recipe_Cost_Analysis!$A$2:$E$6,5,FALSE)'
    style_sheet(plan, {"A": 14, "B": 14, "C": 12, "D": 34, "E": 28, "F": 18})
    for cell in plan["F"][1:]:
        cell.number_format = '#,##0.00" ₽"'
    add_excel_table(plan, f"A1:F{plan.max_row}", "Meal_Weekly_Plan")
    workbook.save(path)


def configure_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.6)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].font.color.rgb = RGBColor.from_string(INK)
    styles["Normal"].paragraph_format.space_after = Pt(6)
    for style_name, size, color in (("Title", 26, NAVY), ("Heading 1", 17, NAVY), ("Heading 2", 12, BLUE)):
        style = styles[style_name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)


def add_metric_table(document: Document, metrics: list[tuple[str, str]]) -> None:
    table = document.add_table(rows=1, cols=len(metrics))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for index, (value, label) in enumerate(metrics):
        cell = table.rows[0].cells[index]
        set_cell_shading(cell, PALE)
        set_cell_text(cell, value, True, BLUE, 15)
        paragraph = cell.add_paragraph(label)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.runs[0]
        run.font.name = "Arial"; run.font.size = Pt(8); run.font.color.rgb = RGBColor.from_string(MUTED)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def write_document(path: Path, result: dict) -> None:
    document = Document()
    configure_document(document)
    title = document.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    title.add_run("Weekly Cafeteria Meal Proposal")
    subtitle = document.add_paragraph()
    run = subtitle.add_run("16–20 марта 2026 · воспроизводимый пакет меню и закупок")
    run.font.name = "Arial"; run.font.size = Pt(11); run.font.color.rgb = RGBColor.from_string(MUTED)
    budget = result["budget_summary"]
    add_metric_table(
        document,
        [
            (f"{budget['Weekly_Total']:,.0f} ₽", "недельный итог"),
            (f"{budget['Average_Daily']:,.2f} ₽", "среднее в день"),
            (f"{budget['Budget_Headroom']:,.0f} ₽", "запас бюджета"),
            (f"{budget['Savings_Vs_Best_Single_Store']:,.0f} ₽", "экономия к лучшему одному магазину"),
        ],
    )

    document.add_heading("Budget Overview", level=1)
    document.add_paragraph(
        f"Недельный budget по методу корзины котировок составляет {budget['Weekly_Total']:,.0f} ₽ при лимите "
        f"{budget['Weekly_Budget']:,.0f} ₽. Средняя стоимость дня — {budget['Average_Daily']:,.2f} ₽, запас — {budget['Budget_Headroom']:,.0f} ₽. "
        f"Лучший сценарий покупки всех корзин в одном магазине — {budget['Best_Single_Store']} "
        f"({budget['Best_Single_Store_Total']:,.0f} ₽); разделение закупок экономит {budget['Savings_Vs_Best_Single_Store']:,.0f} ₽ ({budget['Savings_Pct']:.2f}%)."
    )
    document.add_paragraph(
        "Каждый день остаётся ниже ориентира 2 000 ₽ на Lunch + Dinner. План использует пять разных горячих блюд; каждое повторяется дважды."
    )

    document.add_heading("Menu Highlights", level=1)
    table = document.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = ["Блюдо", "Корзина", "На 4 порции", "С ценой", "Без цены"]
    for cell, value in zip(table.rows[0].cells, headers):
        set_cell_shading(cell, NAVY); set_cell_text(cell, value, True, WHITE, 8)
    for item in result["selected_recipes"]:
        cells = table.add_row().cells
        values = [item["Recipe"], f"{item['Total_Cost']:,.0f} ₽", f"{item['Cost_Per_Serving']:,.2f} ₽", item["Matched_Ingredient_Count"], item["Unpriced_Ingredient_Count"]]
        for cell, value in zip(cells, values):
            set_cell_text(cell, value, False, INK, 8)
    document.add_paragraph(
        "Отбор выполнен не вручную: среди блюд категории «горячее» с тремя и более позициями из прайс-листа взяты пять минимальных корзин."
    )

    document.add_page_break()
    document.add_heading("Procurement Strategy", level=1)
    occurrences: dict[str, dict[str, int]] = {store: defaultdict(int) for store in ("FreshMart", "ValueGrocer", "BulkBarn")}
    selected_by_id = {item["Recipe_ID"]: item for item in result["selected_recipes"]}
    for meal in result["weekly_meal_plan"]:
        for ingredient in selected_by_id[meal["Recipe_ID"]]["Matched_Ingredients"]:
            occurrences[ingredient["Cheapest_Store"]][ingredient["Price_Item"]] += 1
    for store in ("FreshMart", "ValueGrocer", "BulkBarn"):
        document.add_heading(store, level=2)
        items = ", ".join(f"{name} ×{count}" for name, count in sorted(occurrences[store].items()))
        document.add_paragraph(items or "Нет позиций")
    document.add_paragraph(
        "Множитель ×N показывает, сколько рецептных корзин использует одну котируемую единицу. Это не килограммы и не фактический заказ."
    )

    document.add_heading("Limitations", level=1)
    document.add_paragraph(
        "Стоимость — это quoted-basket cost: по одной котируемой единице на каждый разный сопоставленный ингредиент. Единицы веса и штучные позиции не пересчитывались, а цены для отсутствующих в прайс-листе ингредиентов не импутировались."
    )
    document.add_paragraph(
        "Требование белка ≥20 г на порцию: невозможно проверить, поскольку в переданных файлах нет пищевой ценности. Перед утверждением меню нужниы технологические карты или нутрициологический расчёт."
    )
    document.add_paragraph(
        "Источники: grocery_prices.html, all_recipes.json, Meal_Planning_Guidelines.pdf, budget_notes.md. Срез цен: 7 марта 2026."
    )
    document.save(path)


def escape_ical(value: str) -> str:
    return value.replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\n", "\\n")


def fold_ical_line(line: str, width: int = 70) -> list[str]:
    if len(line) <= width:
        return [line]
    pieces = [line[:width]]
    rest = line[width:]
    while rest:
        pieces.append(" " + rest[: width - 1])
        rest = rest[width - 1 :]
    return pieces


def write_calendar(path: Path, result: dict) -> None:
    selected = {item["Recipe_ID"]: item for item in result["selected_recipes"]}
    lunch_rows = [item for item in result["weekly_meal_plan"] if item["Meal"] == "Lunch"]
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//CoworkEvoBench//Meal Planning//RU", "CALSCALE:GREGORIAN"]
    for row in lunch_rows:
        recipe = selected[row["Recipe_ID"]]
        grouped: dict[str, list[str]] = defaultdict(list)
        for ingredient in recipe["Matched_Ingredients"]:
            grouped[ingredient["Cheapest_Store"]].append(ingredient["Price_Item"])
        description_parts = [f"{store}: {', '.join(grouped[store])}" for store in ("FreshMart", "ValueGrocer", "BulkBarn") if grouped[store]]
        if recipe["Unpriced_Ingredients"]:
            description_parts.append("Без котировки: " + ", ".join(recipe["Unpriced_Ingredients"]))
        stamp = row["Date"].replace("-", "")
        event = [
            "BEGIN:VEVENT",
            f"UID:{row['Date']}-{row['Recipe_ID']}@cowork-evo-bench",
            "DTSTAMP:20260307T070000Z",
            f"DTSTART;TZID=Europe/Moscow:{stamp}T080000",
            f"DTEND;TZID=Europe/Moscow:{stamp}T093000",
            "SUMMARY:" + escape_ical("Meal Prep: " + row["Recipe"]),
            "DESCRIPTION:" + escape_ical(" | ".join(description_parts)),
            "END:VEVENT",
        ]
        for line in event:
            lines.extend(fold_ical_line(line))
    lines.append("END:VCALENDAR")
    path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")


def solve(input_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    source = Path(__file__).with_name(PROCESSOR_NAME)
    processor = output_dir / PROCESSOR_NAME
    if source.resolve() != processor.resolve():
        shutil.copy2(source, processor)
    subprocess.run(
        [sys.executable, str(processor), "--prices", str(input_dir / "grocery_prices.html"), "--recipes", str(input_dir / "all_recipes.json"), "--output", str(output_dir / JSON_NAME)],
        check=True, timeout=120,
    )
    result = json.loads((output_dir / JSON_NAME).read_text(encoding="utf-8"))
    write_workbook(output_dir / XLSX_NAME, result)
    write_document(output_dir / DOCX_NAME, result)
    write_calendar(output_dir / ICS_NAME, result)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
