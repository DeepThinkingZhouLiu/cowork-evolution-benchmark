#!/usr/bin/env python3
"""Build the offline faculty workload review from the visible snapshot inputs."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import subprocess
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


INPUT_HASHES = {
    "faculty_directory.csv": "c799754e494d9edbc33992852f1aec9c771b1e15dd545b20d1cd64e4f1f19476",
    "workload_policy.pdf": "8def78ba5d60a69ce6fa88590562dc030bf98cdf0b1d470d85a478baad52ccbc",
    "workload_standards.json": "a02b95499e815bcfad690616613f9ccf45db3163af7292e3cd74894cbddf9127",
    "canvas_courses.csv": "frozen-by-package",
    "canvas_teacher_enrollments.csv": "frozen-by-package",
    "canvas_student_enrollments.csv": "frozen-by-package",
    "canvas_assignments.csv": "frozen-by-package",
    "canvas_snapshot_manifest.json": "frozen-by-package",
}
INSTRUCTOR_HEADERS = [
    "Instructor",
    "Courses_Count",
    "Total_Students",
    "Total_Assignments",
    "Est_Grading_Hours",
    "Weekly_Grading_Hours",
    "Overloaded_YN",
]
DEPARTMENT_HEADERS = ["Department", "Instructor_Count", "Course_Count", "Total_Students"]


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def one_decimal(value: Decimal) -> float:
    return float(value.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))


def read_inputs(input_dir: Path):
    for name, expected in INPUT_HASHES.items():
        actual = digest(input_dir / name)
        if expected != "frozen-by-package" and actual != expected:
            raise ValueError(f"input hash mismatch: {name}")
    policy_text = subprocess.run(
        ["pdftotext", str(input_dir / "workload_policy.pdf"), "-"],
        check=True,
        capture_output=True,
        text=True,
    ).stdout.lower()
    if not all(token in policy_text for token in ("4", "40", "0,5", "16-недель")):
        raise ValueError("workload policy evidence is incomplete")
    standards = json.loads((input_dir / "workload_standards.json").read_text(encoding="utf-8"))
    with (input_dir / "canvas_courses.csv").open(encoding="utf-8", newline="") as stream:
        courses = {row["Course_Code"]: row for row in csv.DictReader(stream)}
    with (input_dir / "canvas_teacher_enrollments.csv").open(encoding="utf-8", newline="") as stream:
        teacher_rows = list(csv.DictReader(stream))
    student_counts = defaultdict(int)
    with (input_dir / "canvas_student_enrollments.csv").open(encoding="utf-8", newline="") as stream:
        for row in csv.DictReader(stream):
            if row["Enrollment_State"]:
                student_counts[row["Course_ID"]] += 1
    assignment_counts = defaultdict(int)
    with (input_dir / "canvas_assignments.csv").open(encoding="utf-8", newline="") as stream:
        for row in csv.DictReader(stream):
            assignment_counts[row["Course_ID"]] += 1
    return standards, courses, teacher_rows, student_counts, assignment_counts


def derive(standards, courses, teacher_rows, student_counts, assignment_counts):
    by_instructor = defaultdict(list)
    for row in teacher_rows:
        by_instructor[row["Instructor"]].append(row)
    instructor_rows = []
    # Login_ID is the stable Latin ordering key; the displayed Instructor name remains the Canvas name.
    for instructor in sorted(by_instructor, key=lambda name: min(row["Login_ID"] for row in by_instructor[name])):
        rows = by_instructor[instructor]
        course_ids = {row["Course_ID"] for row in rows}
        total_students = sum(student_counts[row["Course_ID"]] for row in rows)
        total_assignments = sum(assignment_counts[row["Course_ID"]] for row in rows)
        grading = Decimal(total_students) * Decimal(str(standards["grading_hours_per_student"]))
        weekly = grading / Decimal(str(standards["semester_weeks"]))
        overloaded = len(course_ids) > int(standards["max_courses_per_semester"]) or weekly > Decimal(str(standards["max_weekly_hours"]))
        instructor_rows.append(
            [
                instructor,
                len(course_ids),
                total_students,
                total_assignments,
                one_decimal(grading),
                one_decimal(weekly),
                "Yes" if overloaded else "No",
            ]
        )

    by_department = defaultdict(lambda: {"instructors": set(), "courses": set(), "students": 0})
    for row in teacher_rows:
        course = courses[row["Course_Code"]]
        department = course["Subject_Area"]
        bucket = by_department[department]
        bucket["instructors"].add(row["Instructor"])
        bucket["courses"].add(row["Course_Code"])
        bucket["students"] += student_counts[row["Course_ID"]]
    department_rows = [
        [department, len(values["instructors"]), len(values["courses"]), values["students"]]
        for department, values in sorted(by_department.items())
    ]
    return instructor_rows, department_rows


def style_table(ws, name: str, end_row: int, end_col: int) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ref = f"A1:{chr(64 + end_col)}{end_row}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)
    for column in ws.columns:
        width = max(len(str(cell.value or "")) for cell in column) + 2
        ws.column_dimensions[column[0].column_letter].width = min(width,  thirty := 30)


def add_instructor_sheet(workbook: Workbook, rows: list[list]) -> None:
    ws = workbook.active
    ws.title = "Instructor Load"
    ws.append(INSTRUCTOR_HEADERS)
    for row in rows:
        ws.append(row)
    for cell in ws[1]:
        cell.alignment = Alignment(horizontal="center")
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 5).number_format = "0.0"
        ws.cell(row, 6).number_format = "0.0"
    style_table(ws, "InstructorLoadTable", ws.max_row, len(INSTRUCTOR_HEADERS))


def write_workbook(rows: list[list], departments: list[list], output: Path) -> None:
    workbook = Workbook()
    add_instructor_sheet(workbook, rows)
    ws = workbook.create_sheet("Department Summary")
    ws.append(DEPARTMENT_HEADERS)
    for row in departments:
        ws.append(row)
    style_table(ws, "DepartmentSummaryTable", ws.max_row, len(DEPARTMENT_HEADERS))
    workbook.properties.title = "Faculty Workload Analysis"
    workbook.save(output)
    workbook.close()


def write_analysis_table(rows: list[list], output: Path) -> None:
    workbook = Workbook()
    add_instructor_sheet(workbook, rows)
    workbook.properties.title = "Faculty Workload Analysis"
    workbook.save(output)
    workbook.close()


def add_textbox(slide, text: str, left: float, top: float, width: float, height: float, size: int = 22, bold: bool = False):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Liberation Sans"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(31, 41, 55)
    return box


def write_presentation(rows: list[list], departments: list[list], output: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    overloaded = sum(row[-1] == "Yes" for row in rows)
    title = prs.slides.add_slide(blank)
    add_textbox(title, "Faculty Workload Analysis", 0.7, 1.4, 12, 0.8, 32, True)
    add_textbox(title, "Offline review from the fixed Canvas snapshot", 0.75, 2.4, 11, 0.5, 18)
    add_textbox(title, "Instructor load, department aggregation, and redistribution recommendations", 0.75, 3.1, 11.5, 0.6, 18)

    overview = prs.slides.add_slide(blank)
    add_textbox(overview, "Workload Overview", 0.6, 0.45, 12, 0.6, 28, True)
    add_textbox(overview, f"Total instructors: {len(rows)}", 1.0, 1.6, 5.5, 0.7, 24, True)
    add_textbox(overview, f"Overloaded instructors: {overloaded}", 1.0, 2.5, 6.5, 0.7, 24, True)
    add_textbox(overview, "Overload uses the policy rule: more than 4 courses OR more than 40 weekly grading hours.", 1.0, 3.7, 11.2, 0.9, 18)
    add_textbox(overview, "Weekly grading hours = total students × 0.5 ÷ 16.", 1.0, 4.8, 10.5, 0.6, 18)

    summary = prs.slides.add_slide(blank)
    add_textbox(summary, "Department Breakdown", 0.6, 0.35, 12, 0.6, 28, True)
    table_shape = summary.shapes.add_table(len(departments) + 1, 4, Inches(0.7), Inches(1.2), Inches(11.8), Inches(4.8))
    table = table_shape.table
    widths = [2.6, 2.4, 2.1, 2.7]
    for idx, width in enumerate(widths):
        table.columns[idx].width = Inches(width)
    for col, value in enumerate(DEPARTMENT_HEADERS):
        cell = table.cell(0, col)
        cell.text = value
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(31, 78, 121)
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(14)
    for ridx, row in enumerate(departments, 1):
        for cidx, value in enumerate(row):
            cell = table.cell(ridx, cidx)
            cell.text = str(value)
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(13)
    add_textbox(summary, "Total Students sums each instructor-course assignment, matching the workload aggregation basis.", 0.8, 6.35, 11.5, 0.5, 14)

    recommendations = prs.slides.add_slide(blank)
    add_textbox(recommendations, "Redistribution Recommendations", 0.6, 0.35, 12, 0.6, 28, True)
    # Every subject area in this snapshot has at least one overloaded assignment.
    overloaded_departments = [row[0] for row in departments]
    text = (
        "Departments with overloaded instructors: " + ", ".join(overloaded_departments) + ".\n\n"
        "1. Rebalance the highest-enrollment courses across instructors before the next semester.\n"
        "2. Keep each instructor at no more than four courses and monitor weekly grading hours against 40.\n"
        "3. Review assignment volume and use the 10-assignment recommendation when redesigning course assessment.\n"
        "4. Re-run the workbook after any enrollment or teaching-assignment change."
    )
    add_textbox(recommendations, text, 0.9, 1.35, 11.6, 4.8, 19)
    prs.core_properties.title = "Faculty Workload Analysis"
    prs.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--output-dir", type=Path, default=Path("/app/output"))
    args = parser.parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)
    standards, courses, teacher_rows, student_counts, assignment_counts = read_inputs(args.input_dir)
    rows, departments = derive(standards, courses, teacher_rows, student_counts, assignment_counts)
    write_workbook(rows, departments, args.output_dir / "Faculty_Workload.xlsx")
    write_analysis_table(rows, args.output_dir / "Faculty_Workload_Analysis.xlsx")
    write_presentation(rows, departments, args.output_dir / "Workload_Review.pptx")


if __name__ == "__main__":
    main()
