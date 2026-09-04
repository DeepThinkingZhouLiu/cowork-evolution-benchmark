#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from openpyxl import load_workbook
from docx import Document

KIND = 'prerequisite_check'
INPUT_NAME = 'Thomas_Robinson_Course_List.xlsx'
OUTPUT_NAME = 'prerequisite_check.docx'


def rows(path: Path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    values = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() if value is not None else "" for value in values[0]]
    return headers, values[1:]


def normalized_course(value: object) -> str:
    return "".join(str(value or "").upper().split())


def solve(input_dir: Path, output_dir: Path) -> None:
    headers, data = rows(input_dir / INPUT_NAME)
    index = {name: position for position, name in enumerate(headers)}
    output_dir.mkdir(parents=True, exist_ok=True)
    document = Document()
    if KIND == "roster_counts":
        senior_count = sum(1 for row in data if str(row[index["Academic Standing"]]).strip() == "Senior")
        document.add_heading("CS161 Roster Counts", level=1)
        document.add_paragraph(f"Total students: {len(data)}")
        document.add_paragraph(f"Senior students: {senior_count}")
    elif KIND == "final_pass_count":
        passed = sum(1 for row in data if float(row[index["Final"]]) >= 60)
        document.add_heading("CS161 Final Score Count", level=1)
        document.add_paragraph(f"Students with final score >= 60: {passed}")
    elif KIND == "xihozun_hw2":
        matches = [row for row in data if str(row[index["Last Name"]]).strip() == "Xihozun"]
        if len(matches) != 1:
            raise ValueError(f"expected one Xihozun row, found {len(matches)}")
        document.add_heading("Xihozun HW2 Lookup", level=1)
        document.add_paragraph(f"Xihozun HW2 score: {matches[0][index['HW2']]}")
    elif KIND == "prerequisite_check":
        found = {normalized_course(row[index["Course ID"]]) for row in data}
        required = ["CS101", "CS102", "MATH151", "MATH152"]
        ok = all(course in found for course in required)
        document.add_heading("CS161 Prerequisite Check", level=1)
        document.add_paragraph(f"CS161 prerequisites satisfied: {'Yes' if ok else 'No'}")
        document.add_paragraph("Required courses found: " + ", ".join(required))
    elif KIND == "transfer_eligibility":
        cs_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("CS"))
        document.add_heading("CS Transfer Eligibility", level=1)
        document.add_paragraph(f"CS courses completed: {cs_count}")
        document.add_paragraph(f"Eligible for CS transfer: {'Yes' if cs_count >= 5 else 'No'}")
    elif KIND == "graduate_requirement":
        cs_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("CS"))
        math_count = sum(1 for row in data if normalized_course(row[index["Course ID"]]).startswith("MATH"))
        document.add_heading("Graduate Course Requirement", level=1)
        document.add_paragraph(f"CS courses completed: {cs_count}")
        document.add_paragraph(f"MATH courses completed: {math_count}")
        document.add_paragraph(f"Graduate requirement satisfied: {'Yes' if cs_count >= 8 and math_count >= 5 else 'No'}")
    else:
        raise ValueError(KIND)
    document.save(output_dir / OUTPUT_NAME)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
