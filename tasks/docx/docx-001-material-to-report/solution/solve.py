#!/usr/bin/env python3
"""隐藏评测侧参考解法：根据任务规格生成标准办公交付物。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PptRGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt as PptPt

SPEC = json.loads(r'''{"required_inputs": ["regional_sales.csv", "customer_interviews.md", "report_requirements.json"], "documents": [{"filename": "Q2区域经营复盘.docx", "title": "2026 Q2 区域经营复盘", "subtitle": "经营管理会材料｜商业运营部", "sections": [{"heading": "执行摘要", "paragraphs": ["Q2 总收入为 3,000,000 元，同比增长 8.7%。华东以 1,260,000 元和 18% 的同比增速领先；华北同比下降 4%，应列为优先修复区域。"]}, {"heading": "核心指标", "table": {"headers": ["区域", "Q2收入（元）", "同比增长", "续约率"], "rows": [["华东", "1,260,000", "18%", "94%"], ["华南", "980,000", "9%", "88%"], ["华北", "760,000", "-4%", "82%"], ["合计", "3,000,000", "8.7%", "88%"]]}}, {"heading": "客户反馈与诊断", "bullets": ["华东：增加月度经营复盘，巩固高续约率。", "华南：整合培训材料，建立统一知识库。", "华北：上线支持响应慢，是当前最主要的续约风险。"]}, {"heading": "下一步行动", "table": {"headers": ["优先级", "行动", "负责人", "完成时间"], "rows": [["P0", "建立华北上线支持值班机制", "客户成功负责人", "2026-07-15"], ["P1", "上线统一培训知识库", "交付运营", "2026-07-31"], ["P1", "启动华东月度经营复盘", "区域经理", "2026-07-10"]]}}]}]}''')


def font(run, size=10, bold=None):
    run.font.name = "Noto Sans CJK SC"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def build_docx(output, spec):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(1.8)
    sec.left_margin, sec.right_margin = Cm(2.2), Cm(2.2)
    normal = doc.styles["Normal"]
    normal.font.name = "Noto Sans CJK SC"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
    normal.font.size = Pt(10.5)
    for name, size, color in (("Title", 22, "17365D"), ("Heading 1", 15, "17365D"), ("Heading 2", 12, "2F5597")):
        style = doc.styles[name]
        style.font.name = "Noto Sans CJK SC"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(p.add_run(spec["title"]), 22, True)
    if spec.get("subtitle"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        font(p.add_run(spec["subtitle"]), 10)
    for section in spec["sections"]:
        doc.add_heading(section["heading"], level=section.get("level", 1))
        for text in section.get("paragraphs", []):
            p = doc.add_paragraph()
            font(p.add_run(str(text)), 10)
        for text in section.get("bullets", []):
            p = doc.add_paragraph(style="List Bullet")
            font(p.add_run(str(text)), 10)
        if section.get("table"):
            table_spec = section["table"]
            table = doc.add_table(rows=1, cols=len(table_spec["headers"]))
            table.style = "Table Grid"
            for col, value in enumerate(table_spec["headers"]):
                table.rows[0].cells[col].text = str(value)
                for run in table.rows[0].cells[col].paragraphs[0].runs:
                    font(run, 9, True)
            for values in table_spec["rows"]:
                cells = table.add_row().cells
                for col, value in enumerate(values):
                    cells[col].text = str(value)
                    for run in cells[col].paragraphs[0].runs:
                        font(run, 9)
    for change in spec.get("redlines", []):
        p = doc.add_paragraph()
        font(p.add_run(change.get("label", "修订") + "："), 10, True)
        deletion = OxmlElement("w:del")
        deletion.set(qn("w:author"), "CoworkEvoBench Reviewer")
        deletion.set(qn("w:date"), "2026-08-24T00:00:00Z")
        run = OxmlElement("w:r")
        text = OxmlElement("w:delText")
        text.text = change["old"]
        run.append(text)
        deletion.append(run)
        insertion = OxmlElement("w:ins")
        insertion.set(qn("w:author"), "CoworkEvoBench Reviewer")
        insertion.set(qn("w:date"), "2026-08-24T00:00:00Z")
        run = OxmlElement("w:r")
        text = OxmlElement("w:t")
        text.text = change["new"]
        run.append(text)
        insertion.append(run)
        p._p.append(deletion)
        p._p.append(insertion)
    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    font(footer.add_run("CoworkEvoBench 公开合成样例"), 8)
    doc.save(output)


def build_xlsx(output, spec):
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_spec in spec["sheets"]:
        ws = wb.create_sheet(sheet_spec["name"])
        for row in sheet_spec.get("rows", []):
            ws.append(row)
        for coordinate, value in sheet_spec.get("cells", {}).items():
            ws[coordinate] = value
        for coordinate, value in sheet_spec.get("formulas", {}).items():
            ws[coordinate] = value
        if ws.max_row:
            for cell in ws[1]:
                cell.fill = PatternFill("solid", fgColor="17365D")
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
        for column in ws.columns:
            width = max(len(str(cell.value or "")) for cell in column) + 2
            ws.column_dimensions[column[0].column_letter].width = min(max(width, 11), 34)
        ws.freeze_panes = "A2"
        if ws.max_row > 1 and ws.max_column > 1:
            ws.auto_filter.ref = ws.dimensions
        chart_spec = sheet_spec.get("chart")
        if chart_spec:
            chart = LineChart() if chart_spec.get("type") == "line" else BarChart()
            chart.title = chart_spec.get("title", "指标概览")
            data = Reference(ws, min_col=chart_spec.get("data_col", 2), min_row=1, max_row=ws.max_row)
            categories = Reference(ws, min_col=chart_spec.get("category_col", 1), min_row=2, max_row=ws.max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(categories)
            ws.add_chart(chart, chart_spec.get("anchor", "H2"))
    wb.save(output)


def textbox(slide, text, x, y, w, h, size):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    p = box.text_frame.paragraphs[0]
    p.text = text
    p.font.name = "Noto Sans CJK SC"
    p.font.size = PptPt(size)
    p.font.color.rgb = PptRGBColor(33, 43, 54)
    return box


def build_pptx(output, spec):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for index, slide_spec in enumerate(spec["slides"], 1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = PptRGBColor(246, 248, 251)
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.16))
        band.fill.solid()
        band.fill.fore_color.rgb = PptRGBColor(23, 54, 93)
        band.line.fill.background()
        title = textbox(slide, slide_spec["title"], 0.7, 0.45, 11.8, 0.65, 25)
        title.text_frame.paragraphs[0].font.bold = True
        title.text_frame.paragraphs[0].font.color.rgb = PptRGBColor(23, 54, 93)
        y = 1.35
        for bullet in slide_spec.get("bullets", []):
            textbox(slide, "• " + str(bullet), 0.9, y, 11.4, 0.48, 15)
            y += 0.62
        metrics = slide_spec.get("metrics", [])
        for metric_index, metric in enumerate(metrics):
            x = 0.8 + metric_index * 3.8
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2), Inches(3.4), Inches(1.8))
            card.fill.solid()
            card.fill.fore_color.rgb = PptRGBColor(217, 234, 247)
            card.line.color.rgb = PptRGBColor(91, 155, 213)
            frame = card.text_frame
            frame.clear()
            p = frame.paragraphs[0]
            p.text = str(metric[1])
            p.alignment = PP_ALIGN.CENTER
            p.font.size = PptPt(25)
            p.font.bold = True
            p.font.color.rgb = PptRGBColor(23, 54, 93)
            p = frame.add_paragraph()
            p.text = str(metric[0])
            p.alignment = PP_ALIGN.CENTER
            p.font.size = PptPt(11)
        if slide_spec.get("table"):
            table_spec = slide_spec["table"]
            rows = [table_spec["headers"], *table_spec["rows"]]
            shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(0.8), Inches(1.55), Inches(11.7), Inches(min(4.8, 0.48 * len(rows))))
            for row_index, values in enumerate(rows):
                for col_index, value in enumerate(values):
                    cell = shape.table.cell(row_index, col_index)
                    cell.text = str(value)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = PptRGBColor(23, 54, 93) if row_index == 0 else PptRGBColor(255, 255, 255)
                    p = cell.text_frame.paragraphs[0]
                    p.font.size = PptPt(10)
                    p.font.color.rgb = PptRGBColor(255, 255, 255) if row_index == 0 else PptRGBColor(33, 43, 54)
                    p.font.bold = row_index == 0
        footer = textbox(slide, f"CoworkEvoBench 公开合成样例  |  {index}", 8.5, 7.0, 4.2, 0.25, 8)
        footer.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    prs.save(output)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    missing = [name for name in SPEC["required_inputs"] if not (args.input_dir / name).exists()]
    if missing:
        raise SystemExit(f"缺少输入文件：{missing}")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    for spec in SPEC.get("documents", []):
        build_docx(args.output_dir / spec["filename"], spec)
    for spec in SPEC.get("workbooks", []):
        build_xlsx(args.output_dir / spec["filename"], spec)
    for spec in SPEC.get("presentations", []):
        build_pptx(args.output_dir / spec["filename"], spec)


if __name__ == "__main__":
    main()
