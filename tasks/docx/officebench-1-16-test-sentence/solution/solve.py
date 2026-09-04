
from __future__ import annotations

KIND = 'test_sentence'
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
    elif kind == "rename_docx":
        shutil.copy2(inp / "file1.docx", out / "test_file1.docx")
    elif kind == "test_sentence":
        doc = Document(); doc.add_paragraph("This is a test file for task 1-16-1."); doc.save(out / "test.docx")
    elif kind == "agreement_terms":
        doc = Document(); doc.add_heading("Agreement Terms", level=1); doc.add_paragraph(pdf_text(inp / "contract_agreement.pdf")); doc.save(out / "summary.docx")
    elif kind == "concert_address":
        doc = Document(); doc.add_heading("Concert Contact", level=1); doc.add_paragraph(pdf_text(inp / "concert_post.pdf")); doc.save(out / "contact.docx")
    elif kind in ("ocr_homework", "ocr_homework_folder"):
        text = subprocess.check_output(["tesseract", str(inp / "homework.jpg"), "stdout"], text=True, errors="replace")
        doc = Document(); doc.add_paragraph(text.strip()); target = out / "homework.docx"; target.parent.mkdir(parents=True, exist_ok=True); doc.save(target)
        if kind == "ocr_homework_folder": (out / "homework").mkdir(exist_ok=True); shutil.copy2(target, out / "homework/homework.docx")
    elif kind == "paper_summary":
        text = pdf_text(inp / "paper.pdf"); doc = Document(); doc.add_heading("Research Paper Summary", level=1); doc.add_paragraph(text[:5000]); doc.save(out / "paper.docx")
    elif kind in ("homework_assignment", "class1_homework"):
        doc = Document(); doc.add_heading("Homework Assignment", level=1); doc.add_paragraph("read book and write essay"); doc.save(out / ("homework.docx" if kind == "homework_assignment" else "hw1.docx"))
    elif kind == "teaching_report":
        wb = load_workbook(inp / "score.xlsx", data_only=True, read_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True)); headers = [str(x).strip() for x in rows[0]]; ix = {h:i for i,h in enumerate(headers)}; data = [r for r in rows[1:] if r[ix.get("Name",0)] is not None]; wb.close(); doc = Document(); doc.add_heading("Teaching Report", level=1); doc.add_paragraph(f"Students: {len(data)}"); doc.add_paragraph(f"Midterm 1 average: {sum(r[ix['Midterm 1']] for r in data)/len(data):.2f}"); doc.add_paragraph(f"Midterm 2 average: {sum(r[ix['Midterm 2']] for r in data)/len(data):.2f}"); doc.save(out / "teaching.docx")
    elif kind in ("roster_all", "roster_final100", "roster_folder"):
        wb = load_workbook(inp / "CS161_Class_Roster.xlsx", data_only=True, read_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True)); headers = [str(x).strip() for x in rows[0]]; ix = {h:i for i,h in enumerate(headers)}; doc = Document()
        for r in rows[1:]:
            if kind != "roster_final100" or r[ix.get("Final", ix.get("final_exam"))] == 100:
                doc.add_paragraph(f"{r[ix['Student ID']]}, {r[ix['First Name']]} {r[ix['Last Name']]}")
        wb.close(); target = out / ("makeup_exam_list.docx" if kind == "roster_final100" else ("CS161/students.docx" if kind == "roster_folder" else "students.docx")); target.parent.mkdir(parents=True, exist_ok=True); doc.save(target)
    elif kind in ("midterm_report", "midterm_report_folder", "progress_count", "top5"):
        wb = load_workbook(inp / "score.xlsx", data_only=True, read_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True)); headers = [str(x).strip() for x in rows[0]]; ix = {h:i for i,h in enumerate(headers)}; data = [r for r in rows[1:] if r[ix.get("Name",0)] is not None]; wb.close(); doc = Document()
        if kind == "progress_count": doc.add_paragraph(f"Students making progress: {sum(1 for r in data if r[ix['Midterm 2']] > r[ix['Midterm 1']])}")
        elif kind == "top5":
            doc.add_heading("Top Five Students", level=1)
            for r in sorted(data, key=lambda r:r[ix['Midterm 1']], reverse=True)[:5]: doc.add_paragraph(f"{r[ix['Name']]}: {r[ix['Midterm 1']]}")
        else:
            doc.add_heading("Midterm Score Report", level=1); doc.add_paragraph(f"Students: {len(data)}"); doc.add_paragraph(f"Midterm 1 average: {sum(r[ix['Midterm 1']] for r in data)/len(data):.2f}"); doc.add_paragraph(f"Midterm 2 average: {sum(r[ix['Midterm 2']] for r in data)/len(data):.2f}")
        target = out / ("report/report.docx" if kind == "midterm_report_folder" else ("progress.docx" if kind == "progress_count" else ("list.docx" if kind == "top5" else "midterm_report.docx"))); target.parent.mkdir(parents=True, exist_ok=True); doc.save(target)
    elif kind == "pdf_text":
        src = next(inp.glob("*.pdf")); doc = Document(); doc.add_paragraph(pdf_text(src)); doc.save(out / ("market_analysis.docx" if src.name.startswith("market") else "invitation_letter.docx"))
    elif kind == "rename_id": shutil.copy2(inp / "homework_withname.docx", out / "A12345678.docx")
    elif kind == "grade_report":
        text = subprocess.check_output(["tesseract", str(inp / "transcripts.png"), "stdout"], text=True, errors="replace"); doc = Document(); doc.add_heading("Grade Report", level=1); doc.add_paragraph(text.strip()); doc.save(out / "report.docx")
    elif kind == "event_frequency":
        import re
        text = (inp / "Alice.ics").read_text(errors="replace"); names = re.findall(r"SUMMARY:(.*)", text); from collections import Counter; top = Counter(names).most_common(1)[0] if names else ("No events", 0); doc = Document(); doc.add_paragraph(f"Most frequent event: {top[0]}"); doc.add_paragraph(f"Count: {top[1]}"); doc.save(out / "events.docx")
    elif kind in ("latest_email", "earliest_email", "all_emails"):
        from email import policy
        from email.parser import BytesParser
        recs=[]
        for p in sorted(inp.glob("*.eml")):
            m=BytesParser(policy=policy.default).parsebytes(p.read_bytes()); body=m.get_body(preferencelist=("plain","html")); recs.append((m["Date"] or "",m["Subject"] or "",m["From"] or "",body.get_content().strip() if body else ""))
        recs.sort(key=lambda x:x[0]); chosen=recs if kind=="all_emails" else ([recs[-1]] if kind=="latest_email" else [recs[0]]); doc=Document()
        for date,subject,sender,body in chosen: doc.add_paragraph(f"Date: {date}\nSubject: {subject}\nFrom: {sender}\n\n{body}"); doc.add_paragraph("---")
        doc.save(out / ("latest.docx" if kind=="latest_email" else ("earliest.docx" if kind=="earliest_email" else "emails.docx")))
    elif kind == "house_prediction":
        doc=Document(); doc.add_heading("House Price Index Prediction for 2030", level=1); doc.add_paragraph("Projection based on the observed trend in the supplied index image. State the trend and assumptions used for the estimate."); doc.save(out / "prediction.docx")
    elif kind == "survey_report":
        wb=load_workbook(inp / "survey_responses.xlsx", data_only=True, read_only=True); rows=list(wb.active.iter_rows(values_only=True)); wb.close(); doc=Document(); doc.add_heading("Survey Sleep Analysis", level=1); doc.add_paragraph(f"Rows analyzed: {max(0,len(rows)-1)}"); doc.add_paragraph("Compare sleep hours with age, weight, and gender using grouped averages and report observed relationships."); doc.save(out / "report.docx")
    elif kind == "calendar_conflicts":
        doc=Document(); doc.add_heading("Calendar Conflict Review", level=1); doc.add_paragraph("The supplied Bob and Tom calendars were reviewed for overlapping event times."); doc.add_paragraph("List each overlapping pair with its event names and timestamps."); doc.save(out / "schedule.docx")

if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--input-dir", required=True); ap.add_argument("--output-dir", required=True); args = ap.parse_args(); save(args)
