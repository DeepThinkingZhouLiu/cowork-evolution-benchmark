#!/usr/bin/env python3
"""OfficeBench 3-69 成绩单抽取、GPA 重算与证据边界 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter
from pathlib import Path

from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-3-69-transcript-gpa-audit-package"
RUBRIC_VERSION = "officebench-3-69-transcript-gpa-audit-v0.1"
JUDGE_VERSION = "officebench-3-69-transcript-gpa-audit-deterministic-v0.1"
XLSX_NAME = "Transcript_GPA_Audit.xlsx"
DOCX_NAME = "Transcript_GPA_Audit_Report.docx"
REQUIRED_FILES = [XLSX_NAME, DOCX_NAME]
SOURCE_HASH = "e471b9567cbce1b0335d6d097bc068d4b61d0223d49249ad5dc87095d5d80318"
SHEETS = ["Source_Profile", "Course_Ledger", "Term_Audit", "Summary"]

COURSES = [
    ("2014/2015", "Fall", "Physical Science", "", "A", 1.0),
    ("2014/2015", "Fall", "U.S. History", "", "A", 1.0),
    ("2014/2015", "Winter", "Tennis", "", "A", 1.0),
    ("2014/2015", "Spring", "English Composition", "", "A", 1.0),
    ("2014/2015", "Spring", "Music Appreciation", "", "B", 1.0),
    ("2014/2015", "Summer", "Algebra I", "", "A", 1.0),
    ("2015/2016", "Fall", "American Literature", "", "A", 1.0),
    ("2015/2016", "Fall", "Art Appreciation", "", "A", 1.0),
    ("2015/2016", "Winter", "Soccer", "", "A", 1.0),
    ("2015/2016", "Spring", "Algebra II", "", "B", 1.0),
    ("2015/2016", "Spring", "Biology", "", "A", 1.0),
    ("2015/2016", "Summer", "World History", "", "A", 1.0),
    ("2016/2017", "Fall", "Modern Literature", "", "B", 1.0),
    ("2016/2017", "Fall", "Music Instruction", "", "C", 1.0),
    ("2016/2017", "Winter", "Baseball", "", "A", 1.0),
    ("2016/2017", "Spring", "Geometry", "", "A", 1.0),
    ("2016/2017", "Spring", "Chemistry", "", "C", 1.0),
    ("2016/2017", "Summer", "Trigonometry", "", "A", 1.0),
    ("2017/2018", "Fall", "Geography", "", "A", 1.0),
    ("2017/2018", "Fall", "Physics", "", "A", 1.0),
    ("2017/2018", "Winter", "Economics", "", "A", 1.0),
    ("2017/2018", "Spring", "Drama", "DE", "A", 1.0),
    ("2017/2018", "Spring", "World Literature", "AP", "B", 1.0),
]
TERMS = [
    ("2014/2015", "Fall", 4.00, 4.00, 2.0),
    ("2014/2015", "Winter", 4.00, 4.00, 3.0),
    ("2014/2015", "Spring", 3.50, 3.80, 5.0),
    ("2014/2015", "Summer", 4.00, 3.83, 6.0),
    ("2015/2016", "Fall", 4.00, 3.88, 8.0),
    ("2015/2016", "Winter", 4.00, 3.89, 9.0),
    ("2015/2016", "Spring", 3.50, 3.82, 11.0),
    ("2015/2016", "Summer", 4.00, 3.83, 12.0),
    ("2016/2017", "Fall", 2.50, 3.64, 14.0),
    ("2016/2017", "Winter", 4.00, 3.67, 15.0),
    ("2016/2017", "Spring", 3.00, 3.59, 17.0),
    ("2016/2017", "Summer", 4.00, 3.61, 18.0),
    ("2017/2018", "Fall", 4.00, 3.65, 20.0),
    ("2017/2018", "Winter", 4.00, 3.67, 21.0),
    ("2017/2018", "Spring", 5.00, 3.78, 23.0),
]

PROFILE = [
    ("Student", "Teddy Roosevelt"),
    ("Issue_Date", "October 25, 2018"),
    ("Credits_Earned", 23.0),
    ("Weighted_GPA", 3.78),
    ("Unweighted_GPA", 3.65),
    ("PDF_File_Type", "Image-only"),
    ("Delivered_PDF_Pages", 1),
    ("Printed_Page_Marker", "1 of 2"),
    ("Printed_Total_Pages", 2),
    ("Source_Completeness_Risk", "Potentially incomplete"),
]
LEDGER_HEADERS = [
    "Sequence", "Academic_Year", "Term", "Course", "Program_Tag", "Grade", "Credits",
    "Unweighted_Grade_Points", "Unweighted_Quality_Points",
]
TERM_LEFT_HEADERS = [
    "Academic_Year", "Term", "Course_Count", "Credits", "Unweighted_Quality_Points",
    "Calculated_Unweighted_Term_GPA", "Displayed_Term_GPA",
]
TERM_RIGHT_HEADERS = [
    "Academic_Year", "Term", "Cumulative_Credits", "Cumulative_Unweighted_Points",
    "Calculated_Cumulative_Unweighted_GPA", "Displayed_Cumulative_GPA", "Displayed_vs_Calculated_Delta",
]
SUMMARY_LABELS = [
    "Extracted_Courses", "Extracted_Credits", "Grade_A_Count", "Grade_B_Count", "Grade_C_Count",
    "Unweighted_Quality_Points", "Calculated_Unweighted_GPA", "Rounded_Calculated_Unweighted_GPA",
    "Reported_Unweighted_GPA", "Rounded_Unweighted_Delta", "Reported_Weighted_GPA",
    "Final_Displayed_Cumulative_GPA", "Weighted_vs_Final_Displayed_Delta", "Delivered_PDF_Pages",
    "Printed_Total_Pages", "Source_Completeness_Flag",
]
SUSPICIOUS_MARKERS = ["solution/golden", "judge.py", "expectations.json", "judge-result", "/root/", "/home/", "c:\\"]
PRIVACY_PATTERNS = {
    "street_address": r"1900\s+Pennsylvania\s+Ave",
    "phone": r"\(?202\)?[- )]*456[- ]*1111",
    "date_of_birth": r"(?:Oct(?:ober)?\.?\s*27,?\s*1910|date of birth)",
    "place_of_birth": r"Manhattan,?\s*NY",
}

SPECS = {
    "R001": ("完整交付物与冻结输入", 5, "hurdle", 0.0),
    "R002": ("来源概要与缺页信号", 7, "hurdle", 0.25),
    "R003": ("二十三门课程台账", 13, "hurdle", 0.3),
    "R004": ("课程标签、成绩与学分", 6, "hurdle", 0.4),
    "R005": ("未加权公式与质量分", 12, "hurdle", 0.4),
    "R006": ("十五个学期对账", 12, "hurdle", 0.4),
    "R007": ("汇总与 GPA 重算", 10, "hurdle", 0.45),
    "R008": ("Excel 原生审计结构", 5, "positive", None),
    "R009": ("Word 结构与来源完整性", 7, "hurdle", 0.5),
    "R010": ("Word 学业事实与对账", 8, "hurdle", 0.5),
    "R011": ("加权 GPA 证据边界", 6, "hurdle", 0.55),
    "R012": ("Excel 与 Word 一致", 5, "hurdle", 0.4),
    "R013": ("可渲染与专业排版", 4, "positive", None),
    "R014": ("无课程、成绩、GPA 或公式错误", -30, "penalty", None),
    "R015": ("无跨文件冲突或无来源补全", -25, "penalty", None),
    "R016": ("无缺失、损坏、静态化或调试泄漏", -25, "penalty", None),
    "R017": ("无住址、电话或出生信息泄漏", -15, "penalty", 0.5),
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u00a0", " ").replace("\v", " ")).strip()


def approx(left: object, right: object, tolerance: float = 1e-7) -> bool:
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def grade_points(grade: str) -> int:
    return {"A": 4, "B": 3, "C": 2, "D": 1, "F": 0}[grade]


def source_path() -> Path | None:
    task_root = Path(__file__).resolve().parents[1]
    candidates = [
        Path("/app/input_files/transcripts.pdf"),
        task_root / "data/input_files/transcripts.pdf",
        task_root / "environment/input_files/transcripts.pdf",
    ]
    return next((path for path in candidates if path.is_file()), None)


def validate_source() -> list[str]:
    path = source_path()
    if path is None:
        return ["source_missing"]
    errors = []
    actual_hash = sha256(path)
    if actual_hash != SOURCE_HASH:
        errors.append(f"source_sha256={actual_hash}")
    try:
        proc = subprocess.run(["pdfinfo", str(path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60)
        match = re.search(r"^Pages:\s+(\d+)", proc.stdout, flags=re.MULTILINE)
        if proc.returncode != 0 or not match or int(match.group(1)) != 1:
            errors.append("source_pdf_pages_not_one")
    except Exception as exc:
        errors.append(f"source_pdfinfo:{type(exc).__name__}:{exc}")
    return errors


def inspect_package(path: Path) -> dict[str, object]:
    result: dict[str, object] = {"broken": None, "missing_targets": [], "macros": [], "external": [], "media": []}
    if not path.is_file():
        result["broken"] = "missing"
        return result
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
            result["macros"] = sorted(name for name in names if "vbaProject" in name or name.endswith(".bin"))
            result["media"] = sorted(name for name in names if "/media/" in name)
            for rel_name in [name for name in names if name.endswith(".rels")]:
                root = etree.fromstring(package.read(rel_name))
                for rel in root:
                    target = rel.get("Target", "")
                    if rel.get("TargetMode") == "External":
                        result["external"].append(target)
                        continue
                    if not target or target.startswith("#"):
                        continue
                    if target.startswith("/"):
                        resolved = target.lstrip("/")
                    else:
                        base = Path(rel_name).parent
                        if base.name == "_rels":
                            base = base.parent
                        resolved = str((base / target).as_posix())
                        while "/../" in resolved:
                            resolved = re.sub(r"[^/]+/\.\./", "", resolved, count=1)
                    if resolved not in names:
                        result["missing_targets"].append({"rels": rel_name, "target": target, "resolved": resolved})
    except Exception as exc:
        result["broken"] = f"{type(exc).__name__}:{exc}"
    return result


def recalculate_candidate(path: Path) -> tuple[Path | None, str]:
    root = Path(tempfile.mkdtemp(prefix="transcript-judge-recalc-"))
    input_root = root / "input"
    output_root = root / "output"
    profile = root / "profile"
    input_root.mkdir()
    output_root.mkdir()
    profile.mkdir()
    source = input_root / path.name
    shutil.copy2(path, source)
    try:
        proc = subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "xlsx", "--outdir", str(output_root), str(source)],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
        )
    except Exception as exc:
        return None, f"{type(exc).__name__}:{exc}"
    converted = output_root / path.name
    if proc.returncode != 0 or not converted.is_file():
        return None, proc.stdout[-1000:]
    return converted, proc.stdout[-1000:]


def headers(sheet, columns: list[int], row: int = 1) -> list[str]:
    return [normalize(sheet.cell(row, column).value) for column in columns]


def parse_workbook(path: Path) -> dict[str, object]:
    result: dict[str, object] = {
        "readable": False, "sheetnames": [], "profile_errors": [], "ledger_errors": [], "tag_errors": [],
        "formula_errors": [], "term_errors": [], "summary_errors": [], "native": {},
    }
    try:
        formulas = load_workbook(path, data_only=False)
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}:{exc}"
        return result
    recalculated, log = recalculate_candidate(path)
    result["recalc_log"] = log
    if recalculated is None:
        return result
    try:
        values = load_workbook(recalculated, data_only=True)
    except Exception as exc:
        result["error"] = f"recalc_load:{type(exc).__name__}:{exc}"
        return result
    result["readable"] = True
    result["sheetnames"] = formulas.sheetnames
    if formulas.sheetnames != SHEETS:
        result["profile_errors"].append(f"sheetnames={formulas.sheetnames}")
        return result

    profile = values["Source_Profile"]
    if headers(profile, [1, 2, 3, 4]) != ["Field", "Extracted_Value", "Evidence_Location", "Audit_Note"] or profile.max_row != 11:
        result["profile_errors"].append("profile_structure")
    for row_index, (field, expected) in enumerate(PROFILE, start=2):
        got_field = normalize(profile.cell(row_index, 1).value)
        got_value = profile.cell(row_index, 2).value
        if got_field != field:
            result["profile_errors"].append(f"A{row_index}={got_field}")
        valid = approx(got_value, expected) if isinstance(expected, (int, float)) else normalize(got_value).lower() == str(expected).lower()
        if not valid:
            result["profile_errors"].append(f"B{row_index}={got_value!r}")

    ledger_f = formulas["Course_Ledger"]
    ledger_v = values["Course_Ledger"]
    if headers(ledger_f, list(range(1, 10))) != LEDGER_HEADERS or ledger_f.max_row != 24 or ledger_f.max_column != 9:
        result["ledger_errors"].append("ledger_structure")
    for row_index, expected in enumerate(COURSES, start=2):
        year, term, course, tag, grade, credits = expected
        wanted = [row_index - 1, year, term, course, tag, grade, credits, grade_points(grade), grade_points(grade) * credits]
        for column, target in enumerate(wanted, start=1):
            got = ledger_v.cell(row_index, column).value
            if column in (1, 7, 8, 9):
                valid = approx(got, target)
            else:
                valid = normalize(got).lower() == str(target).lower()
            if not valid:
                bucket = "tag_errors" if column in (5, 6, 7) else "ledger_errors"
                result[bucket].append(f"{ledger_v.cell(row_index, column).coordinate}={got!r}")
        for column in (8, 9):
            cell = ledger_f.cell(row_index, column)
            if cell.data_type != "f":
                result["formula_errors"].append(f"Course_Ledger!{cell.coordinate}:not_formula")
        points_formula = normalize(ledger_f.cell(row_index, 8).value).upper()
        quality_formula = normalize(ledger_f.cell(row_index, 9).value).upper().replace("$", "")
        if not all(token in points_formula for token in (f"F{row_index}", '"A",4', '"B",3', '"C",2', '"D",1', '"F",0')):
            result["formula_errors"].append(f"H{row_index}:wrong_grade_scale")
        if f"G{row_index}*H{row_index}" not in quality_formula:
            result["formula_errors"].append(f"I{row_index}:wrong_quality_source")

    term_f = formulas["Term_Audit"]
    term_v = values["Term_Audit"]
    if (
        headers(term_f, list(range(1, 8))) != TERM_LEFT_HEADERS
        or headers(term_f, list(range(1, 8)), row=18) != TERM_RIGHT_HEADERS
        or term_f.max_row != 33
        or term_f.max_column != 7
    ):
        result["term_errors"].append("term_structure")
    cumulative_credits = 0.0
    cumulative_points = 0.0
    for row_index, (year, term, displayed_term, displayed_cumulative, displayed_credits) in enumerate(TERMS, start=2):
        selected = [course for course in COURSES if course[0] == year and course[1] == term]
        credits = sum(course[5] for course in selected)
        points = sum(grade_points(course[4]) * course[5] for course in selected)
        cumulative_credits += credits
        cumulative_points += points
        wanted_left = [year, term, len(selected), credits, points, points / credits, displayed_term]
        wanted_right = [year, term, cumulative_credits, cumulative_points, cumulative_points / cumulative_credits, displayed_cumulative, displayed_cumulative - cumulative_points / cumulative_credits]
        for column, target in zip(range(1, 8), wanted_left, strict=True):
            got = term_v.cell(row_index, column).value
            valid = approx(got, target) if isinstance(target, (int, float)) else normalize(got) == target
            if not valid:
                result["term_errors"].append(f"{term_v.cell(row_index, column).coordinate}={got!r}")
        cumulative_row = row_index + 17
        for column, target in zip(range(1, 8), wanted_right, strict=True):
            got = term_v.cell(cumulative_row, column).value
            valid = approx(got, target) if isinstance(target, (int, float)) else normalize(got) == target
            if not valid:
                result["term_errors"].append(f"{term_v.cell(cumulative_row, column).coordinate}={got!r}")
        for column in (3, 4, 5, 6):
            if term_f.cell(row_index, column).data_type != "f":
                result["formula_errors"].append(f"Term_Audit!{term_f.cell(row_index, column).coordinate}:not_formula")
        for column in (3, 4, 5, 7):
            if term_f.cell(cumulative_row, column).data_type != "f":
                result["formula_errors"].append(f"Term_Audit!{term_f.cell(cumulative_row, column).coordinate}:not_formula")
        if not approx(displayed_credits, cumulative_credits):
            result["term_errors"].append(f"source_cumulative_credits:{year}/{term}")

    summary_f = formulas["Summary"]
    summary_v = values["Summary"]
    labels = [normalize(summary_v.cell(row, 1).value) for row in range(2, 18)]
    if labels != SUMMARY_LABELS:
        result["summary_errors"].append(f"labels={labels}")
    expected_summary = [
        23, 23.0, 17, 4, 2, 84.0, 84.0 / 23.0, 3.65, 3.65, 0.0, 3.78, 3.78, 0.0, 1, 2, "Potentially incomplete",
    ]
    for row_index, expected in enumerate(expected_summary, start=2):
        if summary_f.cell(row_index, 2).data_type != "f":
            result["formula_errors"].append(f"Summary!B{row_index}:not_formula")
        got = summary_v.cell(row_index, 2).value
        valid = approx(got, expected) if isinstance(expected, (int, float)) else normalize(got).lower() == expected.lower()
        if not valid:
            result["summary_errors"].append(f"B{row_index}={got!r}")

    for sheet in values.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and re.match(r"^#(?:NAME|VALUE|REF|DIV/0|N/A)", cell.value, flags=re.IGNORECASE):
                    result["formula_errors"].append(f"{sheet.title}!{cell.coordinate}={cell.value}")
    result["native"] = {
        "tables": sum(len(sheet.tables) for sheet in formulas.worksheets),
        "charts": sum(len(sheet._charts) for sheet in formulas.worksheets),
        "freeze": all(sheet.freeze_panes == "A2" for sheet in formulas.worksheets),
        "conditional_formats": sum(len(sheet.conditional_formatting) for sheet in formulas.worksheets),
    }
    return result


def word_text(document: Document) -> str:
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def unsupported_claims(text: str) -> list[str]:
    patterns = {
        "invented_missing_page": r"(?:missing|second|page\s*2).{0,35}(?:contains|shows|lists|confirms|reports)",
        "invented_weight_rule": r"(?:the|this|school(?:'s)?)\s+(?:weighted|weighting).{0,30}(?:policy|rule).{0,20}(?:adds|awards|uses|is)",
        "denied_source_gap": r"(?:(?:source|pdf|record).{0,30}(?:is|was)\s+(?:complete|fully complete)|no\s+(?:page|content).{0,20}(?:is|was)\s+missing)",
        "unsupported_authenticity": r"(?:transcript|record).{0,20}(?:is|was)\s+(?:authentic|verified|officially valid)",
        "unsupported_eligibility": r"(?:student|roosevelt).{0,30}(?:is|has been)\s+(?:eligible|graduated|admitted)",
    }
    return [name for name, pattern in patterns.items() if re.search(pattern, text, flags=re.IGNORECASE)]


def privacy_findings(text: str) -> list[str]:
    return [name for name, pattern in PRIVACY_PATTERNS.items() if re.search(pattern, text, flags=re.IGNORECASE)]


def parse_word(path: Path) -> dict[str, object]:
    result: dict[str, object] = {
        "readable": False, "structure_errors": [], "fact_errors": [], "boundary_errors": [],
        "unsupported": [], "privacy": [], "tables": 0, "headings": [],
    }
    try:
        document = Document(path)
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}:{exc}"
        return result
    result["readable"] = True
    result["tables"] = len(document.tables)
    result["headings"] = [normalize(paragraph.text) for paragraph in document.paragraphs if str(paragraph.style.name).startswith("Heading")]
    text = normalize(word_text(document))
    lower = text.lower()
    if "Transcript GPA Audit Report" not in text:
        result["structure_errors"].append("wrong_title")
    required_headings = [
        "Source Completeness", "Reported Academic Profile", "Course Extraction and Recalculation",
        "Term-Level Reconciliation", "Weighted GPA Evidence Boundary", "Audit Conclusion",
    ]
    for heading in required_headings:
        if heading not in result["headings"]:
            result["structure_errors"].append(f"missing_heading:{heading}")
    completeness_tokens = ["image-only PDF", "1 actual page", "1 of 2", "potentially incomplete", "No content from the unprovided page is inferred"]
    for token in completeness_tokens:
        if token.lower() not in lower:
            result["structure_errors"].append(f"missing_completeness:{token}")
    required_facts = [
        "Teddy Roosevelt", "October 25, 2018", "23 courses", "15 terms", "23.00 credits",
        "17 A", "4 B", "2 C", "84.00 unweighted quality points", "84/23 = 3.652173913",
        "reported unweighted GPA", "3.65", "reported weighted GPA", "3.78", "2017/2018 Spring",
        "Drama (DE)", "World Literature (AP)", "3.50", "5.00", "difference of 1.50",
    ]
    for token in required_facts:
        if token.lower() not in lower:
            result["fact_errors"].append(f"missing:{token}")
    for year, term, displayed_term, displayed_cumulative, _ in TERMS:
        period = f"{year} {term}"
        if period.lower() not in lower:
            result["fact_errors"].append(f"missing_period:{period}")
        if f"{displayed_term:.2f}" not in text or f"{displayed_cumulative:.2f}" not in text:
            result["fact_errors"].append(f"missing_term_values:{period}")
    boundary_tokens = [
        "weighting policy", "unprovided second page", "cannot independently reconstruct",
        "does not assume conventional DE/AP bonuses", "reported-but-unverified",
        "not an authenticity, graduation or eligibility determination",
    ]
    for token in boundary_tokens:
        if token.lower() not in lower:
            result["boundary_errors"].append(f"missing:{token}")
    result["unsupported"] = unsupported_claims(text)
    result["privacy"] = privacy_findings(text)
    return result


def render_office(artifact_dir: Path) -> dict[str, object]:
    result: dict[str, object] = {"xlsx_pages": 0, "docx_pages": 0, "errors": [], "logs": []}
    with tempfile.TemporaryDirectory(prefix="transcript-judge-render-") as tmp:
        root = Path(tmp)
        output = root / "output"
        profile = root / "profile"
        output.mkdir()
        profile.mkdir()
        for name in REQUIRED_FILES:
            source = artifact_dir / name
            if not source.is_file():
                result["errors"].append(f"missing:{name}")
                continue
            try:
                proc = subprocess.run(
                    ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", str(output), str(source)],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
                )
            except Exception as exc:
                result["errors"].append(f"{name}:{type(exc).__name__}:{exc}")
                continue
            result["logs"].append(proc.stdout[-1000:])
            pdf = output / f"{Path(name).stem}.pdf"
            if proc.returncode != 0 or not pdf.is_file():
                result["errors"].append(f"{name}:conversion_failed")
                continue
            info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60)
            match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.MULTILINE)
            pages = int(match.group(1)) if match else 0
            extracted = subprocess.run(["pdftotext", str(pdf), "-"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60).stdout
            if name == XLSX_NAME:
                result["xlsx_pages"] = pages
                tokens = ("Teddy Roosevelt", "World Literature", "Displayed_vs_Calculated_Delta", "Potentially incomplete", "Cumulative GPA")
            else:
                result["docx_pages"] = pages
                tokens = ("Transcript GPA Audit Report", "Source Completeness", "Term-Level Reconciliation", "Weighted GPA Evidence Boundary", "reported-but-unverified")
            for token in tokens:
                if token not in extracted:
                    result["errors"].append(f"{name}:render_missing:{token}")
            if re.search(r"#(?:NAME|VALUE|REF|DIV/0|N/A)", extracted, flags=re.IGNORECASE):
                result["errors"].append(f"{name}:formula_error_visible")
    if result["xlsx_pages"] != 4:
        result["errors"].append(f"xlsx_pages={result['xlsx_pages']}")
    if result["docx_pages"] != 3:
        result["errors"].append(f"docx_pages={result['docx_pages']}")
    return result


def suspicious_content(artifact_dir: Path) -> list[str]:
    findings = []
    for path in artifact_dir.iterdir() if artifact_dir.is_dir() else []:
        if not path.is_file():
            continue
        try:
            if path.suffix.lower() in {".md", ".txt"}:
                text = path.read_text(encoding="utf-8", errors="ignore").lower()
            elif path.suffix.lower() in {".xlsx", ".docx"}:
                with zipfile.ZipFile(path) as package:
                    text = " ".join(package.read(name).decode("utf-8", errors="ignore") for name in package.namelist() if name.endswith(".xml")).lower()
            else:
                text = ""
        except Exception:
            continue
        for marker in SUSPICIOUS_MARKERS:
            if marker in text:
                findings.append(f"{path.name}:{marker}")
    return sorted(set(findings))


def artifact_privacy_findings(artifact_dir: Path) -> list[str]:
    """同时审计 Excel 与 Word 的 XML，避免只在报告侧检查隐私泄漏。"""
    findings = []
    for name in REQUIRED_FILES:
        path = artifact_dir / name
        if not path.is_file():
            continue
        try:
            with zipfile.ZipFile(path) as package:
                text = " ".join(
                    package.read(member).decode("utf-8", errors="ignore")
                    for member in package.namelist()
                    if member.endswith(".xml")
                )
        except Exception:
            continue
        findings.extend(f"{name}:{item}" for item in privacy_findings(text))
    return sorted(set(findings))


def criterion(identifier: str, score: float, evidence: object) -> dict[str, object]:
    description, weight, kind, cap = SPECS[identifier]
    raw: dict[str, object] = {"description": description, "weight": weight, "type": kind, "passed": score >= 0.999}
    if cap is not None:
        raw["score_cap_on_failure"] = cap
    return {"criterion_id": identifier, "score": round(float(score), 6), "evidence": evidence, "method": "deterministic", "raw": raw}


def evaluate(artifact_dir: Path) -> dict[str, object]:
    source_errors = validate_source()
    delivered = sorted(path.name for path in artifact_dir.iterdir() if path.is_file()) if artifact_dir.is_dir() else []
    exact_files = delivered == sorted(REQUIRED_FILES)
    packages = {name: inspect_package(artifact_dir / name) for name in REQUIRED_FILES}
    package_ok = all(not info["broken"] and not info["missing_targets"] and not info["macros"] and not info["external"] for info in packages.values())
    workbook = parse_workbook(artifact_dir / XLSX_NAME) if (artifact_dir / XLSX_NAME).is_file() and not source_errors else {
        "readable": False, "profile_errors": ["unavailable"], "ledger_errors": [], "tag_errors": [], "formula_errors": [], "term_errors": [], "summary_errors": [], "native": {},
    }
    word = parse_word(artifact_dir / DOCX_NAME) if (artifact_dir / DOCX_NAME).is_file() and not source_errors else {
        "readable": False, "structure_errors": ["unavailable"], "fact_errors": [], "boundary_errors": [], "unsupported": [], "privacy": [], "tables": 0, "headings": [],
    }
    render = render_office(artifact_dir) if exact_files else {"xlsx_pages": 0, "docx_pages": 0, "errors": ["incomplete"]}
    suspicious = suspicious_content(artifact_dir)
    artifact_privacy = artifact_privacy_findings(artifact_dir)
    profile_ok = workbook.get("readable") and not workbook.get("profile_errors")
    ledger_ok = profile_ok and not workbook.get("ledger_errors")
    tags_ok = ledger_ok and not workbook.get("tag_errors")
    formulas_ok = tags_ok and not workbook.get("formula_errors")
    terms_ok = formulas_ok and not workbook.get("term_errors")
    summary_ok = terms_ok and not workbook.get("summary_errors")
    native = workbook.get("native", {})
    native_ok = native.get("tables", 0) >= 5 and native.get("charts", 0) >= 1 and native.get("freeze") is True and native.get("conditional_formats", 0) >= 5
    word_structure_ok = word.get("readable") and word.get("tables", 0) >= 4 and not word.get("structure_errors") and not packages[DOCX_NAME].get("media")
    word_facts_ok = word_structure_ok and not word.get("fact_errors")
    boundary_ok = word_facts_ok and not word.get("boundary_errors") and not word.get("unsupported")
    privacy_ok = not word.get("privacy") and not artifact_privacy
    cross_ok = summary_ok and word_facts_ok and boundary_ok
    safe_ok = exact_files and package_ok and not suspicious
    render_ok = not render.get("errors")

    results = [
        criterion("R001", 1.0 if exact_files and not source_errors and package_ok else 0.0, {"source_errors": source_errors, "delivered": delivered, "packages": packages}),
        criterion("R002", 1.0 if profile_ok else 0.0, workbook.get("profile_errors")),
        criterion("R003", 1.0 if ledger_ok else 0.0, workbook.get("ledger_errors")),
        criterion("R004", 1.0 if tags_ok else 0.0, workbook.get("tag_errors")),
        criterion("R005", 1.0 if formulas_ok else 0.0, workbook.get("formula_errors")),
        criterion("R006", 1.0 if terms_ok else 0.0, workbook.get("term_errors")),
        criterion("R007", 1.0 if summary_ok else 0.0, workbook.get("summary_errors")),
        criterion("R008", 1.0 if native_ok else 0.0, native),
        criterion("R009", 1.0 if word_structure_ok else 0.0, {"tables": word.get("tables"), "headings": word.get("headings"), "errors": word.get("structure_errors")}),
        criterion("R010", 1.0 if word_facts_ok else 0.0, word.get("fact_errors")),
        criterion("R011", 1.0 if boundary_ok else 0.0, {"boundary_errors": word.get("boundary_errors"), "unsupported": word.get("unsupported")}),
        criterion("R012", 1.0 if cross_ok else 0.0, {"workbook": summary_ok, "word": word_facts_ok, "boundary": boundary_ok}),
        criterion("R013", 1.0 if render_ok and native_ok else 0.0, {"render": render, "native": native}),
        criterion("R014", 1.0 if ledger_ok and tags_ok and formulas_ok and terms_ok and summary_ok else 0.0, {"ledger": ledger_ok, "tags": tags_ok, "formulas": formulas_ok, "terms": terms_ok, "summary": summary_ok}),
        criterion("R015", 1.0 if cross_ok and not word.get("unsupported") else 0.0, {"cross": cross_ok, "unsupported": word.get("unsupported")}),
        criterion("R016", 1.0 if safe_ok and formulas_ok and native_ok else 0.0, {"safe": safe_ok, "formulas": formulas_ok, "native": native_ok, "suspicious": suspicious}),
        criterion("R017", 1.0 if privacy_ok else 0.0, {"word": word.get("privacy"), "all_artifacts": artifact_privacy}),
    ]
    positive_points = sum(SPECS[item["criterion_id"]][1] * float(item["score"]) for item in results if SPECS[item["criterion_id"]][1] > 0)
    penalty_points = sum(SPECS[item["criterion_id"]][1] for item in results if SPECS[item["criterion_id"]][1] < 0 and float(item["score"]) < 0.999)
    reward = max(0.0, min(1.0, (positive_points + penalty_points) / 100.0))
    hurdle_failures = []
    for item in results:
        identifier = item["criterion_id"]
        _, _, kind, cap = SPECS[identifier]
        if kind == "hurdle" and float(item["score"]) < 0.999:
            hurdle_failures.append(identifier)
            reward = min(reward, float(cap))
        elif cap is not None and float(item["score"]) < 0.999:
            reward = min(reward, float(cap))
    reward = round(reward, 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [sha256(artifact_dir / name) for name in REQUIRED_FILES if (artifact_dir / name).is_file()],
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
