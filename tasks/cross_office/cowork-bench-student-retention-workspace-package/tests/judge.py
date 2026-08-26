#!/usr/bin/env python3
"""从 Canvas 源行独立重算真值的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-student-retention-workspace-package"
RUBRIC_VERSION = "cowork-bench-student-retention-workspace-package-v0.1"
JUDGE_VERSION = "cowork-bench-student-retention-workspace-package-v0.1"
PASS_THRESHOLD = 0.8
OUTPUT_NAMES = [
    "risk_scorer.py",
    "risk_analysis_results.json",
    "Student_Risk_Analysis.xlsx",
    "Intervention_Plan.docx",
    "Student_Risk_Tracker.md",
    "advisor_email.txt",
]
RISK_ORDER = ["High", "Medium", "Low"]
COURSE_IDS = [16, 17]
OVERVIEW_HEADERS = ["course_id", "course_name", "enrollment_count", "avg_score", "pass_rate"]
DISTRIBUTION_HEADERS = ["risk_level", "student_count", "pct"]
AT_RISK_HEADERS = ["course_name", "high_risk_count", "medium_risk_count", "low_risk_count"]
INTERVENTION_HEADERS = ["risk_level", "action", "timeline", "responsible"]
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.3,
    "R003": 0.4,
    "R004": 0.45,
    "R005": 0.45,
    "R006": 0.45,
    "R007": 0.5,
    "R008": 0.55,
    "R009": 0.55,
    "R010": 0.55,
    "R011": 0.6,
    "R014": 0.7,
}
CRITERION_SPECS = {
    "R001": ("六个指定交付物均存在且可打开", 6, "hurdle"),
    "R002": ("评分器在原始输入上可回放且等于源行重算真值", 8, "hurdle"),
    "R003": ("High 学生分数变异后完整结果正确联动，不是固定答案", 9, "hurdle"),
    "R004": ("JSON 课程概览的 enrollment、graded、均分与 pass rate 正确", 7, "hurdle"),
    "R005": ("JSON 风险分布、分课程人数、干预和 Tracker 摘要正确", 7, "hurdle"),
    "R006": ("Excel 四表的表头、行数与全部数据正确", 8, "hurdle"),
    "R007": ("Excel 风险百分比为引用人数的原生公式且重算值正确", 6, "hurdle"),
    "R008": ("Word 五个章节、核心事实、行动和数据边界完整正确", 8, "hurdle"),
    "R009": ("Markdown 跟踪页的 H1、表头和两行课程事实正确", 6, "hurdle"),
    "R010": ("离线邮件的收件人、主题、High 人数和三级行动正确", 6, "hurdle"),
    "R011": ("JSON、Excel、Word、Markdown 与邮件无事实冲突", 6, "hurdle"),
    "R012": ("Excel 与 Word 保持原生可编辑结构并达到专业交付要求", 4, "positive"),
    "R013": ("Excel 与 Word 可由 LibreOffice 渲染且核心文本可提取", 3, "positive"),
    "R014": ("报告和邮件没有无来源姓名、因果、趋势或效果断言", 5, "hurdle"),
    "R015": ("任一来源派生数值与 Canvas 输入重算真值冲突", -15, "penalty"),
    "R016": ("评分器不可回放或不随分数变异输入改变", -14, "penalty"),
    "R017": ("交付物缺失、损坏、截图化、无法重算或无法渲染", -14, "penalty"),
    "R018": ("核心事实跨文件冲突或出现输入不支持的断言", -18, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str = "") -> dict[str, Any]:
    description, weight, kind = CRITERION_SPECS[identifier]
    return {
        "criterion_id": identifier,
        "description": description,
        "weight": weight,
        "type": kind,
        "passed": bool(passed),
        "score": 1.0 if passed else 0.0,
        "evidence": evidence,
        "method": "deterministic",
    }


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def classify(value: float, high_cutoff: float, low_cutoff: float) -> str:
    if value < high_cutoff:
        return "High"
    if value < low_cutoff:
        return "Medium"
    return "Low"


def parse_cutoffs(model: dict[str, Any]) -> tuple[float, float]:
    high_text = model["risk_classification"]["high_risk"]["threshold"]
    medium_text = model["risk_classification"]["medium_risk"]["threshold"]
    high_values = re.findall(r"-?\d+(?:\.\d+)?", high_text)
    medium_values = re.findall(r"-?\d+(?:\.\d+)?", medium_text)
    if not high_values or len(medium_values) < 2:
        raise ValueError("scoring_model.json 风险阈值不可解析")
    high_cutoff = float(high_values[0])
    low_cutoff = round(float(medium_values[1]) + 0.01, 2)
    if float(medium_values[0]) != high_cutoff:
        raise ValueError("风险阈值不连续")
    return high_cutoff, low_cutoff


def derive_truth(input_dir: Path, include_student_details: bool = False) -> dict[str, Any]:
    courses = {
        int(row["course_id"]): row
        for row in read_csv(input_dir / "canvas_courses.csv")
        if int(row["course_id"]) in COURSE_IDS
    }
    if sorted(courses) != COURSE_IDS:
        raise ValueError(f"课程集合错误：{sorted(courses)}")
    enrollments = read_csv(input_dir / "canvas_enrollments.csv")
    enrollment_counts = Counter(
        int(row["course_id"])
        for row in enrollments
        if int(row["course_id"]) in COURSE_IDS and "student" in row["type"].lower()
    )
    assignments = read_csv(input_dir / "canvas_assignments.csv")
    assignment_course = {
        int(row["assignment_id"]): int(row["course_id"])
        for row in assignments
        if int(row["course_id"]) in COURSE_IDS
    }
    model = json.loads((input_dir / "scoring_model.json").read_text(encoding="utf-8"))
    high_cutoff, low_cutoff = parse_cutoffs(model)

    student_scores: dict[tuple[int, int], list[float]] = defaultdict(list)
    course_scores: dict[int, list[float]] = defaultdict(list)
    for row in read_csv(input_dir / "canvas_submissions.csv"):
        assignment_id = int(row["assignment_id"])
        score_text = row["score"].strip()
        if assignment_id not in assignment_course or score_text == "":
            continue
        course_id = assignment_course[assignment_id]
        score = float(score_text)
        student_scores[(course_id, int(row["user_id"]))].append(score)
        course_scores[course_id].append(score)

    per_course: dict[int, Counter[str]] = {course_id: Counter() for course_id in COURSE_IDS}
    student_details: list[dict[str, Any]] = []
    for (course_id, user_id), scores in sorted(student_scores.items()):
        average = sum(scores) / len(scores)
        risk = classify(average, high_cutoff, low_cutoff)
        per_course[course_id][risk] += 1
        student_details.append({
            "course_id": course_id,
            "user_id": user_id,
            "average": average,
            "risk_level": risk,
        })

    overview: list[dict[str, Any]] = []
    at_risk: list[dict[str, Any]] = []
    tracker: list[dict[str, Any]] = []
    for course_id in COURSE_IDS:
        risks = per_course[course_id]
        graded_students = sum(risks.values())
        average = round(sum(course_scores[course_id]) / len(course_scores[course_id]), 2)
        pass_rate = round(100.0 * (risks["Medium"] + risks["Low"]) / graded_students, 2)
        predominant = min(RISK_ORDER, key=lambda level: (-risks[level], RISK_ORDER.index(level)))
        overview.append({
            "course_id": course_id,
            "course_name": courses[course_id]["course_name"],
            "enrollment_count": enrollment_counts[course_id],
            "graded_student_count": graded_students,
            "graded_submission_count": len(course_scores[course_id]),
            "avg_score": average,
            "pass_rate": pass_rate,
        })
        at_risk.append({
            "course_name": courses[course_id]["course_name"],
            "high_risk_count": risks["High"],
            "medium_risk_count": risks["Medium"],
            "low_risk_count": risks["Low"],
        })
        tracker.append({
            "Student Course": courses[course_id]["course_name"],
            "Risk Level": predominant,
            "Student Count": graded_students,
            "Average Score": average,
            "Pass Rate": pass_rate,
        })

    overall: Counter[str] = Counter()
    for risks in per_course.values():
        overall.update(risks)
    total = sum(overall.values())
    distribution = [
        {
            "risk_level": level,
            "student_count": overall[level],
            "pct": round(100.0 * overall[level] / total, 2),
        }
        for level in RISK_ORDER
    ]
    actions = model["intervention_actions"]
    intervention = [
        {"risk_level": level, **actions[level.lower()]}
        for level in RISK_ORDER
    ]
    result: dict[str, Any] = {
        "course_overview": overview,
        "risk_distribution": distribution,
        "at_risk_students": at_risk,
        "intervention_plan": intervention,
        "student_risk_tracker": tracker,
    }
    if include_student_details:
        result["_student_details"] = student_details
    return result


def close(actual: Any, expected: Any, tolerance: float = 0.011) -> bool:
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return False


def payload_equal(actual: Any, expected: Any) -> bool:
    if isinstance(expected, dict):
        return isinstance(actual, dict) and set(actual) == set(expected) and all(
            payload_equal(actual[key], expected[key]) for key in expected
        )
    if isinstance(expected, list):
        return isinstance(actual, list) and len(actual) == len(expected) and all(
            payload_equal(left, right) for left, right in zip(actual, expected)
        )
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return close(actual, expected)
    return actual == expected


def public_truth(truth: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in truth.items() if not key.startswith("_")}


def run_scorer(script: Path, input_dir: Path) -> tuple[dict[str, Any] | None, str]:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-retention-scorer-") as directory:
        output = Path(directory) / "result.json"
        command = [
            sys.executable,
            str(script.resolve()),
            "--courses", str((input_dir / "canvas_courses.csv").resolve()),
            "--enrollments", str((input_dir / "canvas_enrollments.csv").resolve()),
            "--assignments", str((input_dir / "canvas_assignments.csv").resolve()),
            "--submissions", str((input_dir / "canvas_submissions.csv").resolve()),
            "--model", str((input_dir / "scoring_model.json").resolve()),
            "--output", str(output),
        ]
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
                check=False,
            )
        except Exception as error:
            return None, f"执行异常：{error}"
        if completed.returncode != 0 or not output.is_file():
            return None, f"退出={completed.returncode}，输出存在={output.is_file()}，日志={completed.stdout[-1000:]}"
        try:
            return json.loads(output.read_text(encoding="utf-8")), completed.stdout[-1000:]
        except Exception as error:
            return None, f"JSON 无法读取：{error}"


def build_mutated_inputs(source: Path, target: Path, truth: dict[str, Any]) -> tuple[int, int, int]:
    shutil.copytree(source, target)
    high_students = [
        row for row in truth["_student_details"] if row["risk_level"] == "High"
    ]
    selected = min(high_students, key=lambda row: (row["user_id"], row["course_id"]))
    selected_course = int(selected["course_id"])
    selected_user = int(selected["user_id"])
    assignment_course = {
        int(row["assignment_id"]): int(row["course_id"])
        for row in read_csv(target / "canvas_assignments.csv")
    }
    submission_path = target / "canvas_submissions.csv"
    rows = read_csv(submission_path)
    fields = list(rows[0])
    changed = 0
    for row in rows:
        if (
            int(row["user_id"]) == selected_user
            and assignment_course.get(int(row["assignment_id"])) == selected_course
            and row["score"].strip()
        ):
            row["score"] = "100.00"
            changed += 1
    with submission_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    return selected_course, selected_user, changed


def sheet_rows(sheet: Any, headers: list[str], columns: int | None = None) -> list[dict[str, Any]] | None:
    column_count = columns or len(headers)
    actual = [sheet.cell(1, index).value for index in range(1, column_count + 1)]
    if actual != headers:
        return None
    rows: list[dict[str, Any]] = []
    for row_number in range(2, sheet.max_row + 1):
        values = [sheet.cell(row_number, index).value for index in range(1, column_count + 1)]
        if all(value in (None, "") for value in values):
            continue
        rows.append(dict(zip(headers, values)))
    return rows


def read_excel(path: Path) -> tuple[dict[str, Any] | None, dict[str, Any], str]:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        return None, {}, str(error)
    expected_sheets = ["Course_Overview", "Risk_Distribution", "At_Risk_Students", "Intervention_Plan"]
    formatting = {
        "sheet_names": workbook.sheetnames,
        "tables": sum(len(sheet.tables) for sheet in workbook.worksheets),
        "freezes": {sheet.title: str(sheet.freeze_panes or "") for sheet in workbook.worksheets},
        "images": sum(len(sheet._images) for sheet in workbook.worksheets),
        "header_bold": all(
            all(cell.font.bold for cell in sheet[1] if cell.value is not None)
            for sheet in workbook.worksheets
        ),
        "column_widths": all(
            any((dimension.width or 0) > 12 for dimension in sheet.column_dimensions.values())
            for sheet in workbook.worksheets
        ),
    }
    if workbook.sheetnames != expected_sheets:
        workbook.close()
        return None, formatting, f"工作表={workbook.sheetnames}"
    overview = sheet_rows(workbook["Course_Overview"], OVERVIEW_HEADERS)
    distribution = sheet_rows(workbook["Risk_Distribution"], DISTRIBUTION_HEADERS)
    at_risk = sheet_rows(workbook["At_Risk_Students"], AT_RISK_HEADERS)
    intervention = sheet_rows(workbook["Intervention_Plan"], INTERVENTION_HEADERS)
    formulas = [workbook["Risk_Distribution"][cell].value for cell in ("C2", "C3", "C4")]
    workbook.close()
    if any(rows is None for rows in (overview, distribution, at_risk, intervention)):
        return None, formatting, "表头不匹配"
    return {
        "overview": overview,
        "distribution": distribution,
        "at_risk": at_risk,
        "intervention": intervention,
        "formulas": formulas,
    }, formatting, ""


def recalculate_excel(path: Path, target: Path) -> tuple[list[Any] | None, str]:
    source_dir = target / "source"
    output_dir = target / "output"
    source_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    copied = source_dir / path.name
    shutil.copy2(path, copied)
    env = os.environ.copy()
    env["HOME"] = str(target / "lo-home")
    env["NO_PROXY"] = "localhost,127.0.0.1,::1"
    try:
        completed = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "xlsx", "--outdir", str(output_dir), str(copied)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=120,
            check=False,
            env=env,
        )
    except Exception as error:
        return None, str(error)
    result = output_dir / path.name
    if completed.returncode != 0 or not result.is_file():
        return None, f"退出={completed.returncode}，文件={result.is_file()}，日志={completed.stdout[-500:]}"
    try:
        workbook = load_workbook(result, data_only=True)
        values = [workbook["Risk_Distribution"][cell].value for cell in ("C2", "C3", "C4")]
        workbook.close()
        return values, completed.stdout[-500:]
    except Exception as error:
        return None, str(error)


def excel_checks(payload: dict[str, Any] | None, truth: dict[str, Any]) -> tuple[bool, list[str]]:
    if not payload:
        return False, ["Excel 不可读取"]
    errors: list[str] = []
    expected_overview = [
        {key: row[key] for key in OVERVIEW_HEADERS}
        for row in truth["course_overview"]
    ]
    if not payload_equal(payload["overview"], expected_overview):
        errors.append("Course_Overview 不匹配")
    actual_distribution = [
        {"risk_level": row["risk_level"], "student_count": row["student_count"]}
        for row in payload["distribution"]
    ]
    expected_distribution = [
        {"risk_level": row["risk_level"], "student_count": row["student_count"]}
        for row in truth["risk_distribution"]
    ]
    if not payload_equal(actual_distribution, expected_distribution):
        errors.append("Risk_Distribution 人数不匹配")
    if not payload_equal(payload["at_risk"], truth["at_risk_students"]):
        errors.append("At_Risk_Students 不匹配")
    if not payload_equal(payload["intervention"], truth["intervention_plan"]):
        errors.append("Intervention_Plan 不匹配")
    return not errors, errors


def docx_payload(path: Path) -> tuple[str, list[str], int, int]:
    document = Document(path)
    pieces: list[str] = []
    headings: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        pieces.append(text)
        if paragraph.style.name == "Title" or paragraph.style.name.lower().startswith("heading"):
            headings.append(text)
    for table in document.tables:
        for row in table.rows:
            pieces.extend(cell.text.strip() for cell in row.cells if cell.text.strip())
    with zipfile.ZipFile(path) as archive:
        media = sum(
            1 for name in archive.namelist()
            if name.startswith("word/media/") and not name.endswith("/")
        )
    return "\n".join(pieces), headings, len(document.tables), media


def contains_number(text: str, number: Any, percent: bool = False) -> bool:
    value = float(number)
    variants = {f"{value:.2f}", f"{value:,.2f}"}
    if value.is_integer():
        variants.update({str(int(value)), f"{int(value):,}"})
    if percent:
        return any(f"{variant}%" in text for variant in variants)
    return any(variant in text for variant in variants)


def word_checks(text: str, headings: list[str], truth: dict[str, Any]) -> tuple[bool, list[str]]:
    errors: list[str] = []
    required = [
        "Student Retention Intervention Plan",
        "Executive Summary",
        "Course Overview",
        "Risk Distribution",
        "Recommended Interventions",
        "Method & Scope",
    ]
    for title in required:
        if not any(title.lower() == heading.lower() for heading in headings):
            errors.append(f"缺标题 {title}")
    for row in truth["course_overview"]:
        if row["course_name"] not in text:
            errors.append(f"缺课程 {row['course_id']}")
        for key in ("enrollment_count", "graded_student_count", "graded_submission_count", "avg_score"):
            if not contains_number(text, row[key]):
                errors.append(f"课程 {row['course_id']} 缺 {key}={row[key]}")
        if not contains_number(text, row["pass_rate"], percent=True):
            errors.append(f"课程 {row['course_id']} 缺 pass_rate={row['pass_rate']}")
    for row in truth["risk_distribution"]:
        if row["risk_level"] not in text or not contains_number(text, row["student_count"]):
            errors.append(f"缺风险人数 {row['risk_level']}")
        if not contains_number(text, row["pct"], percent=True):
            errors.append(f"缺风险百分比 {row['risk_level']}")
    for row in truth["intervention_plan"]:
        for key in ("action", "timeline", "responsible"):
            if row[key] not in text:
                errors.append(f"缺干预 {row['risk_level']}:{key}")
    if not all(token in text for token in ("High <60", "Medium 60", "Low ≥75")):
        errors.append("缺方法阈值")
    if not any(token in text.lower() for token in ("не содержат причин", "does not contain causes", "no causal")):
        errors.append("缺无因果边界")
    return not errors, errors


def tracker_checks(text: str, truth: dict[str, Any]) -> tuple[bool, list[str]]:
    errors: list[str] = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines or lines[0] != "# Student Risk Tracker":
        errors.append("H1 错误")
    table_lines = [line for line in lines if line.startswith("|") and line.endswith("|")]
    if len(table_lines) < 4:
        errors.append("Markdown 表格不足两行")
        return False, errors
    header = [cell.strip() for cell in table_lines[0].strip("|").split("|")]
    expected_header = ["Student Course", "Risk Level", "Student Count", "Average Score", "Pass Rate"]
    if header != expected_header:
        errors.append(f"表头错误={header}")
    data_lines = table_lines[2:]
    if len(data_lines) != 2:
        errors.append(f"数据行={len(data_lines)}")
    else:
        for line, expected in zip(data_lines, truth["student_risk_tracker"]):
            values = [cell.strip() for cell in line.strip("|").split("|")]
            wanted = [
                expected["Student Course"],
                expected["Risk Level"],
                str(expected["Student Count"]),
                f"{expected['Average Score']:.2f}",
                f"{expected['Pass Rate']:.2f}%",
            ]
            if values != wanted:
                errors.append(f"行不匹配：{values} != {wanted}")
    return not errors, errors


def email_checks(text: str, truth: dict[str, Any]) -> tuple[bool, list[str]]:
    errors: list[str] = []
    if not re.search(r"^To:\s*academic_advisors@university\.edu\s*$", text, re.MULTILINE):
        errors.append("收件人错误")
    if not re.search(r"^Subject:\s*Student Retention Risk Analysis - Action Required\s*$", text, re.MULTILINE):
        errors.append("主题错误")
    high_count = next(row["student_count"] for row in truth["risk_distribution"] if row["risk_level"] == "High")
    if not contains_number(text, high_count):
        errors.append("High 总人数错误")
    for row in truth["intervention_plan"]:
        if row["action"] not in text or row["timeline"] not in text:
            errors.append(f"缺 {row['risk_level']} 行动或时限")
    return not errors, errors


def unsupported_claims(text: str, truth: dict[str, Any]) -> list[str]:
    findings: list[str] = []
    allowed_percentages = {60.0, 74.99, 75.0}
    allowed_percentages.update(float(row["pct"]) for row in truth["risk_distribution"])
    allowed_percentages.update(float(row["pass_rate"]) for row in truth["course_overview"])
    for match in re.finditer(r"(?<![\d.])(\d+(?:\.\d+)?)\s*%", text):
        value = float(match.group(1))
        if not any(abs(value - allowed) <= 0.011 for allowed in allowed_percentages):
            findings.append(f"无来源百分比 {match.group(1)}%")
    lowered = text.lower()
    patterns = [
        (r"(?:intervention|advising|tutoring)[^.\n]{0,90}(?:increased|improved|raised|reduced)[^.\n]{0,90}\d", "无来源干预效果"),
        (r"(?:increased|improved|raised|reduced)[^.\n]{0,90}(?:pass|retention|score|risk)", "无来源效果"),
        (r"(?:year[- ]over[- ]year|\byoy\b|upward trend|downward trend)", "无来源趋势"),
        (r"(?:high risk|low scores?)[^.\n]{0,50}(?:caused by|because of|due to)", "无来源因果"),
        (r"(?:student|learner)\s+(?:john|jane|ivan|petr|maria)\s+[a-z]+", "无来源学生姓名"),
        (r"(?:повысил|улучшил|снизил)[^.\n]{0,80}\d", "无来源干预效果"),
        (r"(?:годовой тренд|рост по сравнению|снижение по сравнению)", "无来源趋势"),
        (r"(?:высокий риск|низкая успеваемость)[^.\n]{0,60}(?:вызван|из-за)", "无来源因果"),
    ]
    for pattern, label in patterns:
        if re.search(pattern, lowered):
            findings.append(label)
    return sorted(set(findings))


def render_office(paths: list[Path], target: Path) -> tuple[dict[str, dict[str, Any]], str]:
    rendered: dict[str, dict[str, Any]] = {}
    logs: list[str] = []
    target.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["HOME"] = str(target / "lo-home")
    env["NO_PROXY"] = "localhost,127.0.0.1,::1"
    for path in paths:
        try:
            completed = subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", str(target), str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                timeout=120,
                check=False,
                env=env,
            )
        except Exception as error:
            logs.append(f"{path.name}: {error}")
            continue
        pdf = target / f"{path.stem}.pdf"
        logs.append(f"{path.name}: rc={completed.returncode}, pdf={pdf.is_file()}, {completed.stdout[-240:]}")
        if completed.returncode != 0 or not pdf.is_file():
            continue
        info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
        page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        extracted = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)
        rendered[path.name] = {
            "pages": int(page_match.group(1)) if page_match else 0,
            "text": extracted.stdout,
            "pdf": str(pdf),
        }
    return rendered, " | ".join(logs)


def build_result(results: list[dict[str, Any]], artifact_paths: list[Path]) -> dict[str, Any]:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付缺失或前置门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
    digests = [
        hashlib.sha256(path.read_bytes()).hexdigest()
        for path in artifact_paths
        if path.is_file()
    ] or [hashlib.sha256(b"").hexdigest()]
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append({
            "criterion_id": item["criterion_id"],
            "score": score,
            "evidence": item["evidence"],
            "method": item["method"],
            "raw": {
                "description": item["description"],
                "weight": item["weight"],
                "type": item["type"],
                "passed": item["passed"],
            },
        })
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests,
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdle_failures,
    }


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = task_root / "data" / "input_files"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results: list[dict[str, Any]] = []
    missing = [name for name, path in paths.items() if not path.is_file()]
    structure_ok = not missing
    results.append(criterion("R001", structure_ok, f"缺失={missing}"))
    if missing:
        results.append(criterion("R017", True, f"缺失={missing}"))
        return build_result(results, list(paths.values()))

    truth_with_details = derive_truth(input_dir, include_student_details=True)
    truth = public_truth(truth_with_details)
    replay, replay_log = run_scorer(paths["risk_scorer.py"], input_dir)
    replay_ok = replay is not None and payload_equal(replay, truth)
    results.append(criterion("R002", replay_ok, f"匹配={replay_ok}，日志={replay_log}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-retention-mutation-") as directory:
        mutated_input = Path(directory) / "inputs"
        selected_course, selected_user, changed_rows = build_mutated_inputs(input_dir, mutated_input, truth_with_details)
        mutated_truth = derive_truth(mutated_input)
        mutated_replay, mutation_log = run_scorer(paths["risk_scorer.py"], mutated_input)
    mutation_changed = (
        changed_rows > 0
        and mutated_truth["course_overview"] != truth["course_overview"]
        and mutated_truth["risk_distribution"] != truth["risk_distribution"]
        and mutated_truth["at_risk_students"] != truth["at_risk_students"]
    )
    mutation_ok = mutated_replay is not None and payload_equal(mutated_replay, mutated_truth) and mutation_changed
    results.append(criterion(
        "R003",
        mutation_ok,
        f"course={selected_course}，user={selected_user}，改动行={changed_rows}，匹配={mutated_replay is not None and payload_equal(mutated_replay, mutated_truth)}，真值变化={mutation_changed}，日志={mutation_log}",
    ))

    try:
        delivered_json = json.loads(paths["risk_analysis_results.json"].read_text(encoding="utf-8"))
        overview_json_ok = isinstance(delivered_json, dict) and payload_equal(delivered_json.get("course_overview"), truth["course_overview"])
        remaining_json_ok = isinstance(delivered_json, dict) and all(
            payload_equal(delivered_json.get(key), truth[key])
            for key in ("risk_distribution", "at_risk_students", "intervention_plan", "student_risk_tracker")
        ) and set(delivered_json) == set(truth)
        json_error = ""
    except Exception as error:
        delivered_json = None
        overview_json_ok = False
        remaining_json_ok = False
        json_error = str(error)
    results.append(criterion("R004", overview_json_ok, f"匹配={overview_json_ok}，错误={json_error}"))
    results.append(criterion("R005", remaining_json_ok, f"匹配={remaining_json_ok}，错误={json_error}"))

    excel_payload, formatting, excel_error = read_excel(paths["Student_Risk_Analysis.xlsx"])
    excel_ok, excel_errors = excel_checks(excel_payload, truth)
    results.append(criterion("R006", excel_ok, f"错误={excel_errors or excel_error}"))
    formulas = excel_payload.get("formulas", []) if excel_payload else []
    formula_shape_ok = len(formulas) == 3 and all(
        isinstance(value, str)
        and value.startswith("=")
        and re.search(r"B[2-4]", value, re.IGNORECASE)
        and "SUM" in value.upper()
        for value in formulas
    )
    with tempfile.TemporaryDirectory(prefix="cowork-evo-retention-recalc-") as directory:
        recalculated, recalc_log = recalculate_excel(paths["Student_Risk_Analysis.xlsx"], Path(directory))
    expected_pct = [row["pct"] for row in truth["risk_distribution"]]
    recalc_ok = recalculated is not None and len(recalculated) == 3 and all(
        close(actual, expected) for actual, expected in zip(recalculated, expected_pct)
    )
    formula_ok = formula_shape_ok and recalc_ok
    results.append(criterion("R007", formula_ok, f"公式={formulas}，重算={recalculated}，日志={recalc_log}"))

    try:
        word_text, word_headings, word_tables, word_media = docx_payload(paths["Intervention_Plan.docx"])
        word_ok, word_errors = word_checks(word_text, word_headings, truth)
    except Exception as error:
        word_text, word_headings, word_tables, word_media = "", [], 0, 0
        word_ok, word_errors = False, [str(error)]
    results.append(criterion("R008", word_ok, f"错误={word_errors}，标题={len(word_headings)}，表={word_tables}"))

    try:
        tracker_text = paths["Student_Risk_Tracker.md"].read_text(encoding="utf-8")
        tracker_ok, tracker_errors = tracker_checks(tracker_text, truth)
    except Exception as error:
        tracker_text = ""
        tracker_ok, tracker_errors = False, [str(error)]
    results.append(criterion("R009", tracker_ok, f"错误={tracker_errors}"))

    try:
        email_text = paths["advisor_email.txt"].read_text(encoding="utf-8")
        email_ok, email_errors = email_checks(email_text, truth)
    except Exception as error:
        email_text = ""
        email_ok, email_errors = False, [str(error)]
    results.append(criterion("R010", email_ok, f"错误={email_errors}"))

    json_ok = overview_json_ok and remaining_json_ok
    cross_ok = json_ok and excel_ok and word_ok and tracker_ok and email_ok
    results.append(criterion("R011", cross_ok, f"JSON={json_ok}，Excel={excel_ok}，Word={word_ok}，Markdown={tracker_ok}，Email={email_ok}"))

    native_ok = (
        formatting.get("images") == 0
        and formatting.get("tables", 0) >= 4
        and formatting.get("freezes") == {
            "Course_Overview": "A2",
            "Risk_Distribution": "A2",
            "At_Risk_Students": "A2",
            "Intervention_Plan": "A2",
        }
        and formatting.get("header_bold")
        and formatting.get("column_widths")
        and word_media == 0
        and word_tables >= 2
        and len(word_headings) >= 6
    )
    results.append(criterion("R012", native_ok, json.dumps(formatting | {"word_tables": word_tables, "word_media": word_media, "word_headings": len(word_headings)}, ensure_ascii=False)))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-retention-render-") as directory:
        rendered, render_log = render_office(
            [paths["Student_Risk_Analysis.xlsx"], paths["Intervention_Plan.docx"]],
            Path(directory),
        )
    xlsx_render = rendered.get("Student_Risk_Analysis.xlsx", {})
    docx_render = rendered.get("Intervention_Plan.docx", {})
    render_ok = (
        xlsx_render.get("pages", 0) >= 4
        and all(token in xlsx_render.get("text", "") for token in ("course_id", "risk_level", "high_risk_count", "responsible"))
        and docx_render.get("pages", 0) >= 2
        and all(token in docx_render.get("text", "") for token in ("Executive Summary", "Risk Distribution", "Method & Scope"))
    )
    results.append(criterion("R013", render_ok, f"页数=xlsx:{xlsx_render.get('pages', 0)},docx:{docx_render.get('pages', 0)}，日志={render_log}"))

    unsupported = unsupported_claims(word_text + "\n" + email_text, truth)
    no_unsupported = not unsupported
    results.append(criterion("R014", no_unsupported, f"发现={unsupported}"))

    source_error = not (replay_ok and json_ok and excel_ok and word_ok and tracker_ok and email_ok)
    script_error = not (replay_ok and mutation_ok)
    damaged_error = not (structure_ok and excel_payload is not None and bool(word_text) and formula_ok and render_ok)
    conflict_or_unsupported = not cross_ok or bool(unsupported)
    results.append(criterion("R015", source_error, f"评分器={replay_ok}，JSON={json_ok}，Excel={excel_ok}，Word={word_ok}，Markdown={tracker_ok}，Email={email_ok}"))
    results.append(criterion("R016", script_error, f"原始回放={replay_ok}，变异回放={mutation_ok}"))
    results.append(criterion("R017", damaged_error, f"结构={structure_ok}，Excel={excel_payload is not None}，Word={bool(word_text)}，公式={formula_ok}，渲染={render_ok}"))
    results.append(criterion("R018", conflict_or_unsupported, f"跨文件一致={cross_ok}，无来源断言={unsupported}"))
    return build_result(results, list(paths.values()))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve())
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
