
from __future__ import annotations

KIND = 'concert_contact'
import argparse, shutil, subprocess
from collections import defaultdict
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

def pdf_text(path: Path) -> str:
    return subprocess.check_output(["pdftotext", "-layout", str(path), "-"], text=True)

def save(args):
    inp, out = Path(args.input_dir), Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)
    kind = KIND
    if kind == "house_creak":
        doc = Document(inp / "House_creak.docx")
        doc.add_paragraph("This is a test sentence for task 1-15-0.")
        doc.save(out / "house_creak.docx")
    elif kind == "project_report":
        doc = Document(inp / "project_report.docx")
        doc.add_paragraph((inp / "reference.txt").read_text(encoding="utf-8"))
        doc.save(out / "project_report_edited.docx")
    elif kind == "random_paragraph":
        doc = Document(); doc.add_paragraph((inp / "random_paragraph.txt").read_text(encoding="utf-8")); doc.save(out / "random_paragraph.docx")
    elif kind == "paper_metadata":
        doc = Document(); doc.add_heading("OS-Copilot: Towards Generalist Computer Agents with Self-Improvement", 0)
        doc.add_paragraph("Authors: Zhiyong Wu; Chengcheng Han; Zichen Ding; Zhenmin Weng; Zhoumianze Liu; Shunyu Yao; Tao Yu; Lingpeng Kong")
        doc.save(out / "paper.docx")
    elif kind == "concert_contact":
        doc = Document(); doc.add_heading("Concert Contact", 0)
        for label, value in [("Artist", "John Smith"), ("Date", "July 15, 2024"), ("Time", "7:00 PM"), ("Venue", "Music Hall"), ("Address", "123 Main Street, Cityville"), ("Tickets", "www.example.com/concert")]:
            doc.add_paragraph(f"{label}: {value}")
        doc.save(out / "contact.docx")
    elif kind == "shopping_list":
        wb = load_workbook(inp / "shopping_list.xlsx", data_only=True, read_only=True); ws = wb.active
        doc = Document(); table = doc.add_table(rows=1, cols=2); table.style = "Table Grid"
        table.rows[0].cells[0].text, table.rows[0].cells[1].text = "Name", "Amount"
        for name, amount in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            if name is not None:
                c = table.add_row().cells; c[0].text, c[1].text = str(name), str(amount)
        wb.close(); doc.save(out / "shopping_list.docx")
    elif kind == "cs_major_list":
        wb = load_workbook(inp / "CS161_Class_Roster.xlsx", data_only=True, read_only=True); ws = wb.active
        headers = [str(x).strip() if x is not None else "" for x in next(ws.iter_rows(values_only=True))]; ix = {h:i for i,h in enumerate(headers)}
        doc = Document()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if str(row[ix["Major"]]).strip() == "CS": doc.add_paragraph(f"{row[ix['Student ID']]}, {row[ix['First Name']]} {row[ix['Last Name']]}")
        wb.close(); doc.save(out / "cs_class_2024.docx")
    elif kind == "bill_courses":
        wb = load_workbook(inp / "Bill_course_lists.xlsx", data_only=True, read_only=True); ws = wb.active
        headers = [str(x).strip() if x is not None else "" for x in next(ws.iter_rows(values_only=True))]; ci = headers.index("Course")
        doc = Document()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[ci] is not None: doc.add_paragraph(str(row[ci]))
        wb.close(); doc.save(out / "classes_taken.docx")

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--input-dir", required=True); ap.add_argument("--output-dir", required=True); args = ap.parse_args(); save(args)
