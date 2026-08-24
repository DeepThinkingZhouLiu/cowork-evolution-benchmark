#!/usr/bin/env python3
"""ppt-clickhouse-executive 的逐题离线参考解法。

仅处理本题从 Cowork-Bench SQL dump 抽取的 Q4 2024 订单和客户维表，
按上游 Query 聚合地区与客户分群，交付 Excel、可编辑 PPTX 和其 PDF 归档。
这不是批量制题器。
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import shutil
import subprocess
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Cm, Pt


ORDERS_NAME = "q4_orders_2024.csv"
CUSTOMERS_NAME = "customer_dimension.csv"
XLSX_NAME = "Q4_Executive_Analysis.xlsx"
PPTX_NAME = "Q4_Executive_Briefing.pptx"
PDF_NAME = "Q4_Executive_Briefing.pdf"

ORDERS_SHA256 = "432161d192c7239e634350125d9f5347cdc4b0d70a7f264a463cbfb351fb4ce3"
CUSTOMERS_SHA256 = "0f0e68e115c96312a1fa0188912689271d6a15e0765f915fa15e078d7a34a575"

PAGE_W_CM = 33.867
PAGE_H_CM = 19.05

NAVY = "12304A"
NAVY_2 = "1B4563"
TEAL = "1AA6A6"
CYAN = "5FD0D5"
BLUE = "3D79A8"
GOLD = "E5AA45"
INK = "1B2633"
MUTED = "66788A"
LIGHT = "F4F7FA"
WHITE = "FFFFFF"
PALE = "E7EEF4"
GRID = "D3DFE8"

REGION_ACCENTS = ["1AA6A6", "3D79A8", "6A79C5", "D4855A", "2E9974"]


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_source(input_dir: Path) -> tuple[list[dict], list[dict]]:
    orders_path = input_dir / ORDERS_NAME
    customers_path = input_dir / CUSTOMERS_NAME
    for path, expected in ((orders_path, ORDERS_SHA256), (customers_path, CUSTOMERS_SHA256)):
        if not path.is_file():
            raise FileNotFoundError(f"缺少输入文件：{path}")
        actual = sha256(path)
        if actual != expected:
            raise ValueError(f"{path.name} SHA256 不匹配：{actual}")

    with customers_path.open(encoding="utf-8", newline="") as stream:
        customers = list(csv.DictReader(stream))
    with orders_path.open(encoding="utf-8", newline="") as stream:
        orders = list(csv.DictReader(stream))
    if len(customers) != 1451 or len(orders) != 2579:
        raise ValueError(f"输入行数不符：customers={len(customers)}, orders={len(orders)}")
    return customers, orders


def aggregate(customers: list[dict], orders: list[dict]) -> tuple[list[dict], list[dict], int, float]:
    customer_by_id = {row["CUSTOMER_ID"]: row for row in customers}
    region_accumulator: dict[str, list[float]] = defaultdict(lambda: [0, 0.0])
    segment_accumulator: dict[str, list[float]] = defaultdict(lambda: [0, 0.0])

    for order in orders:
        customer = customer_by_id.get(order["CUSTOMER_ID"])
        if customer is None:
            raise ValueError(f"订单 {order['ORDER_ID']} 的客户 {order['CUSTOMER_ID']} 不在维表中")
        if not ("2024-10-01" <= order["ORDER_DATE"] < "2025-01-01"):
            raise ValueError(f"非 Q4 2024 订单：{order}")
        amount = float(order["TOTAL_AMOUNT"])
        for accumulator, key in (
            (region_accumulator, customer["REGION"]),
            (segment_accumulator, customer["SEGMENT"]),
        ):
            accumulator[key][0] += 1
            accumulator[key][1] += amount

    regions = []
    for region, (orders_count, revenue) in sorted(region_accumulator.items()):
        revenue = round(revenue, 2)
        regions.append(
            {
                "name": region,
                "orders": int(orders_count),
                "revenue": revenue,
                "avg_order_value": round(revenue / orders_count, 2),
            }
        )

    total_orders = sum(row["orders"] for row in regions)
    total_revenue = round(sum(row["revenue"] for row in regions), 2)
    segments = []
    for segment, (orders_count, revenue) in sorted(segment_accumulator.items()):
        revenue = round(revenue, 2)
        segments.append(
            {
                "name": segment,
                "orders": int(orders_count),
                "revenue": revenue,
                "revenue_share_pct": round(revenue / total_revenue * 100, 1),
            }
        )

    expected_total = (2579, 411857.18)
    if (total_orders, total_revenue) != expected_total or len(regions) != 5 or len(segments) != 4:
        raise ValueError(
            f"聚合结果不符合上游真值："
            f"regions={len(regions)}, segments={len(segments)}, total={(total_orders, total_revenue)}"
        )
    return regions, segments, total_orders, total_revenue


def style_worksheet(ws, widths: list[float], currency_columns: tuple[int, ...], pct_columns: tuple[int, ...] = ()) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 27
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="Aptos", size=11, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color=GRID)
    for row_index in range(2, ws.max_row + 1):
        ws.row_dimensions[row_index].height = 23
        for cell in ws[row_index]:
            cell.fill = PatternFill("solid", fgColor=WHITE if row_index % 2 == 0 else LIGHT)
            cell.font = Font(name="Aptos", size=10, color=INK)
            cell.alignment = Alignment(horizontal="left" if cell.column == 1 else "right", vertical="center")
            cell.border = Border(bottom=thin)
    for column_index, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + column_index)].width = width
    for column_index in currency_columns:
        for cell in ws.iter_cols(min_col=column_index, max_col=column_index, min_row=2):
            for item in cell:
                item.number_format = '$#,##0.00'
    for column_index in pct_columns:
        for cell in ws.iter_cols(min_col=column_index, max_col=column_index, min_row=2):
            for item in cell:
                item.number_format = '0.0'
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = "Q4 2024 Executive Analysis"
    ws.oddFooter.right.text = "Page &P of &N"


def build_excel(regions: list[dict], segments: list[dict], output_path: Path) -> None:
    workbook = Workbook()
    regional = workbook.active
    regional.title = "Regional Summary"
    regional.append(["Region", "Orders", "Revenue", "Avg_Order_Value"])
    for row in regions:
        regional.append([row["name"], row["orders"], row["revenue"], row["avg_order_value"]])
    style_worksheet(regional, [36, 14, 18, 20], (3, 4))
    table = Table(displayName="RegionalSummaryTable", ref=regional.dimensions)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False)
    regional.add_table(table)

    segment = workbook.create_sheet("Segment Analysis")
    segment.append(["Segment", "Orders", "Revenue", "Revenue_Share_Pct"])
    for row in segments:
        segment.append([row["name"], row["orders"], row["revenue"], row["revenue_share_pct"]])
    style_worksheet(segment, [34, 14, 18, 22], (3,), (4,))
    table = Table(displayName="SegmentAnalysisTable", ref=segment.dimensions)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False)
    segment.add_table(table)

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    workbook.save(output_path)


def set_fill(shape, color: str) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)


def set_outline(shape, color: str | None, width_pt: float = 1.0) -> None:
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(color)
        shape.line.width = Pt(width_pt)


def add_shape(slide, shape_type, name: str, left: float, top: float, width: float, height: float, fill: str, line: str | None = None, line_width: float = 1.0):
    shape = slide.shapes.add_shape(shape_type, Cm(left), Cm(top), Cm(width), Cm(height))
    shape.name = name
    set_fill(shape, fill)
    set_outline(shape, line, line_width)
    return shape


def add_text(
    slide,
    name: str,
    text: str,
    left: float,
    top: float,
    width: float,
    height: float,
    *,
    size: float,
    color: str = INK,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    vertical=MSO_ANCHOR.MIDDLE,
    font_name: str = "Liberation Sans",
):
    shape = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    shape.name = name
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = vertical
    frame.margin_left = frame.margin_right = Cm(0.08)
    frame.margin_top = frame.margin_bottom = Cm(0.05)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = paragraph.space_after = Pt(0)
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return shape


def add_header(slide, title: str, page_number: int, accent: str = TEAL) -> None:
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, "Header background", 0, 0, PAGE_W_CM, 2.05, NAVY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, "Header accent", 0, 0, 0.28, 2.05, accent)
    add_text(slide, "Slide title", title, 1.05, 0.30, 27.0, 1.22, size=25, color=WHITE, bold=True)
    add_text(slide, "Page number", f"{page_number:02d}", 30.7, 0.38, 1.6, 0.95, size=13, color=CYAN, bold=True, align=PP_ALIGN.RIGHT)


def add_metric_card(slide, name: str, label: str, value: str, left: float, top: float, width: float, accent: str) -> None:
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, f"{name} card", left, top, width, 4.15, WHITE, GRID, 1.0)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, f"{name} accent", left, top, 0.22, 4.15, accent)
    add_text(slide, f"{name} label", label.upper(), left + 0.65, top + 0.50, width - 1.2, 0.68, size=10.5, color=MUTED, bold=True)
    add_text(slide, f"{name} value", value, left + 0.65, top + 1.30, width - 1.2, 1.70, size=27, color=NAVY, bold=True)


def build_title_slide(slide) -> None:
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = rgb(NAVY)
    # 可编辑的几何装饰，不使用图片背景。
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, "Decorative teal circle", 25.8, -3.2, 12.0, 12.0, NAVY_2)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, "Decorative cyan circle", 29.2, 11.2, 7.4, 7.4, TEAL)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, "Title accent line", 2.35, 4.05, 2.25, 0.16, GOLD)
    add_text(slide, "Eyebrow", "QUARTERLY LEADERSHIP REVIEW", 2.35, 2.70, 13.0, 0.70, size=11.5, color=CYAN, bold=True)
    add_text(slide, "Main title", "Q4 2024 Sales\nExecutive Briefing", 2.35, 4.55, 24.0, 4.25, size=35, color=WHITE, bold=True)
    add_text(slide, "Subtitle", "October – December 2024", 2.35, 9.20, 15.0, 0.85, size=17, color=PALE)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Period chip", 2.35, 12.55, 7.25, 1.30, NAVY_2, TEAL, 1.2)
    add_text(slide, "Period chip text", "REGION & SEGMENT VIEW", 2.70, 12.72, 6.55, 0.92, size=11.5, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_text(slide, "Source line", "Source: Cowork-Bench SALES_DW snapshot", 2.35, 16.65, 19.0, 0.65, size=9.5, color="A8BBC9")


def add_region_comparison(slide, regions: list[dict], current_index: int) -> None:
    left, top, width, height = 20.65, 4.02, 11.05, 9.72
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Regional comparison panel", left, top, width, height, WHITE, GRID, 1.0)
    add_text(slide, "Regional comparison heading", "Q4 revenue by region", left + 0.55, top + 0.42, width - 1.1, 0.65, size=12, color=NAVY, bold=True)
    maximum = max(row["revenue"] for row in regions)
    for index, row in enumerate(regions):
        y = top + 1.45 + index * 1.46
        color = REGION_ACCENTS[current_index] if index == current_index else "C7D3DD"
        add_text(slide, f"Comparison label {index + 1}", row["name"], left + 0.55, y, 4.50, 0.55, size=8.8, color=INK if index == current_index else MUTED, bold=index == current_index)
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, f"Comparison track {index + 1}", left + 5.05, y + 0.05, 4.20, 0.38, PALE)
        bar_width = 4.20 * row["revenue"] / maximum
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, f"Comparison bar {index + 1}", left + 5.05, y + 0.05, bar_width, 0.38, color)
        add_text(slide, f"Comparison value {index + 1}", f"${row['revenue']:,.0f}", left + 9.25, y - 0.03, 1.25, 0.55, size=8.2, color=MUTED, align=PP_ALIGN.RIGHT)


def build_region_slide(slide, row: dict, regions: list[dict], index: int, total_revenue: float) -> None:
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = rgb(LIGHT)
    accent = REGION_ACCENTS[index]
    add_header(slide, row["name"], index + 2, accent)
    add_text(slide, "Section label", "REGIONAL PERFORMANCE", 1.10, 2.65, 8.0, 0.65, size=10.5, color=accent, bold=True)
    add_metric_card(slide, "Orders", "Orders", f"{row['orders']:,}", 1.10, 4.02, 8.75, accent)
    add_metric_card(slide, "Revenue", "Revenue", f"${row['revenue']:,.2f}", 10.30, 4.02, 9.55, accent)
    share = row["revenue"] / total_revenue * 100
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Revenue share callout", 1.10, 9.05, 18.75, 4.69, WHITE, GRID, 1.0)
    add_text(slide, "Revenue share label", "SHARE OF Q4 REVENUE", 1.75, 9.62, 8.0, 0.65, size=10.5, color=MUTED, bold=True)
    add_text(slide, "Revenue share value", f"{share:.1f}%", 1.75, 10.52, 5.0, 1.35, size=29, color=accent, bold=True)
    add_text(slide, "Average order label", "Average order value", 8.15, 10.10, 5.0, 0.65, size=10.5, color=MUTED)
    add_text(slide, "Average order value", f"${row['avg_order_value']:,.2f}", 8.15, 10.78, 6.0, 1.05, size=21, color=NAVY, bold=True)
    add_text(slide, "Metric definition", "Revenue ÷ orders", 8.15, 11.88, 5.0, 0.55, size=9.2, color=MUTED)
    add_region_comparison(slide, regions, index)
    add_text(slide, "Footer", "Q4 2024 • Executive briefing", 1.10, 17.62, 12.0, 0.55, size=9.0, color=MUTED)


def build_summary_slide(slide, regions: list[dict], total_orders: int, total_revenue: float) -> None:
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = rgb(NAVY)
    add_text(slide, "Summary eyebrow", "QUARTER CLOSE", 2.15, 1.15, 8.0, 0.65, size=11.0, color=CYAN, bold=True)
    add_text(slide, "Summary title", "Q4 2024 Summary", 2.15, 2.05, 18.0, 1.30, size=30, color=WHITE, bold=True)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, "Summary accent line", 2.15, 3.72, 2.0, 0.14, GOLD)

    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Total orders card", 2.15, 4.75, 8.55, 4.25, NAVY_2, "41627B", 1.0)
    add_text(slide, "Total orders label", "TOTAL ORDERS", 2.75, 5.25, 7.2, 0.65, size=10.5, color=CYAN, bold=True)
    add_text(slide, "Total orders value", f"{total_orders:,}", 2.75, 6.10, 7.2, 1.55, size=32, color=WHITE, bold=True)

    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Total revenue card", 11.20, 4.75, 10.35, 4.25, NAVY_2, "41627B", 1.0)
    add_text(slide, "Total revenue label", "TOTAL REVENUE", 11.80, 5.25, 9.0, 0.65, size=10.5, color=CYAN, bold=True)
    add_text(slide, "Total revenue value", f"${total_revenue:,.2f}", 11.80, 6.10, 9.0, 1.55, size=29, color=WHITE, bold=True)

    top_region = max(regions, key=lambda row: row["revenue"])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, "Top region card", 22.05, 4.75, 9.65, 4.25, TEAL, None)
    add_text(slide, "Top region label", "HIGHEST REVENUE REGION", 22.65, 5.25, 8.35, 0.65, size=9.7, color=WHITE, bold=True)
    add_text(slide, "Top region name", top_region["name"], 22.65, 6.02, 8.35, 1.15, size=17, color=WHITE, bold=True)
    add_text(slide, "Top region revenue", f"${top_region['revenue']:,.2f}", 22.65, 7.28, 8.35, 0.68, size=12.5, color=WHITE)

    add_text(slide, "Regional detail heading", "REGIONAL DETAIL", 2.15, 10.20, 8.0, 0.62, size=10.5, color=CYAN, bold=True)
    row_top = 11.15
    for index, row in enumerate(regions):
        y = row_top + index * 1.03
        if index % 2 == 0:
            add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, f"Summary row fill {index + 1}", 2.15, y - 0.10, 29.55, 0.92, NAVY_2)
        add_text(slide, f"Summary region {index + 1}", row["name"], 2.55, y, 14.0, 0.58, size=9.8, color=WHITE, bold=index == regions.index(top_region))
        add_text(slide, f"Summary orders {index + 1}", f"{row['orders']:,} orders", 17.20, y, 5.0, 0.58, size=9.8, color=PALE, align=PP_ALIGN.RIGHT)
        add_text(slide, f"Summary revenue {index + 1}", f"${row['revenue']:,.2f}", 24.15, y, 6.85, 0.58, size=9.8, color=WHITE, bold=True, align=PP_ALIGN.RIGHT)
    add_text(slide, "Summary source", "Source: Q4 orders joined to customer region and segment", 2.15, 17.35, 22.0, 0.55, size=9.0, color="A8BBC9")


def build_presentation(regions: list[dict], total_orders: int, total_revenue: float, output_path: Path) -> None:
    presentation = Presentation()
    presentation.slide_width = Cm(PAGE_W_CM)
    presentation.slide_height = Cm(PAGE_H_CM)
    blank = presentation.slide_layouts[6]

    title_slide = presentation.slides.add_slide(blank)
    build_title_slide(title_slide)
    for index, row in enumerate(regions):
        slide = presentation.slides.add_slide(blank)
        build_region_slide(slide, row, regions, index, total_revenue)
    summary = presentation.slides.add_slide(blank)
    build_summary_slide(summary, regions, total_orders, total_revenue)
    presentation.save(output_path)


def export_pdf(pptx_path: Path, output_dir: Path) -> Path:
    executable = shutil.which("libreoffice")
    if executable is None:
        raise RuntimeError("缺少 libreoffice，无法生成上游 Query 要求的 PDF 归档")
    completed = subprocess.run(
        [executable, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(pptx_path)],
        capture_output=True,
        text=True,
        timeout=180,
        check=False,
    )
    pdf_path = output_dir / PDF_NAME
    if completed.returncode != 0 or not pdf_path.is_file():
        raise RuntimeError((completed.stderr or completed.stdout or "LibreOffice PDF 转换失败")[-1200:])
    return pdf_path


def solve(input_dir: Path, output_dir: Path) -> list[Path]:
    customers, orders = load_source(input_dir)
    regions, segments, total_orders, total_revenue = aggregate(customers, orders)
    output_dir.mkdir(parents=True, exist_ok=True)

    xlsx_path = output_dir / XLSX_NAME
    pptx_path = output_dir / PPTX_NAME
    build_excel(regions, segments, xlsx_path)
    build_presentation(regions, total_orders, total_revenue, pptx_path)
    pdf_path = export_pdf(pptx_path, output_dir)
    return [xlsx_path, pptx_path, pdf_path]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    for path in solve(args.input_dir, args.output_dir):
        print(path)


if __name__ == "__main__":
    main()
