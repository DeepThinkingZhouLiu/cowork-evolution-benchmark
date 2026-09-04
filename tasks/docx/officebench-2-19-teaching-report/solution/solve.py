#!/usr/bin/env python3
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


def read_rows(path: Path):
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    index = {name: position for position, name in enumerate(headers)}
    return [(str(row[index["Name"]]).strip(), float(row[index["Score"]])) for row in rows[1:] if row[index["Name"]] is not None]


def band(score: float) -> str:
    return "A" if score >= 90 else "B" if score >= 80 else "C" if score >= 70 else "D" if score >= 60 else "F"


def solve(input_dir: Path, output_dir: Path) -> None:
    records = read_rows(input_dir / "score.xlsx")
    grouped = defaultdict(list)
    for name, score in records:
        grouped[name].append(score)
    document = Document()
    document.add_heading("Teaching Report", level=0)
    document.add_paragraph("Computed from score.xlsx; repeated names are consolidated by student.")
    document.add_heading("Class Overview", level=1)
    distribution = Counter(band(score) for _, score in records)
    document.add_paragraph(f"Grade records: {len(records)} | Students: {len(grouped)} | Class average: {sum(score for _, score in records) / len(records):.2f}")
    document.add_paragraph("Grade distribution: " + ", ".join(f"{key}={distribution[key]}" for key in "ABCDF"))
    document.add_heading("Student Summary", level=1)
    table = document.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    for cell, label in zip(table.rows[0].cells, ["Student", "Records", "Average", "Highest", "Lowest"]):
        cell.text = label
    for name in sorted(grouped):
        scores = grouped[name]
        cells = table.add_row().cells
        cells[0].text = name
        cells[1].text = str(len(scores))
        cells[2].text = f"{sum(scores) / len(scores):.2f}"
        cells[3].text = f"{max(scores):.0f}"
        cells[4].text = f"{min(scores):.0f}"
    document.add_heading("Teaching Priorities", level=1)
    support = [name for name, scores in grouped.items() if sum(scores) / len(scores) < 70]
    praise = [name for name, scores in grouped.items() if sum(scores) / len(scores) >= 90]
    document.add_paragraph("Priority support: " + (", ".join(support) if support else "None") + ".")
    document.add_paragraph("High-performing students: " + (", ".join(praise) if praise else "None") + ".")
    output_dir.mkdir(parents=True, exist_ok=True)
    document.save(output_dir / "teaching.docx")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    solve(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
