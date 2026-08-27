#!/usr/bin/env python3
"""从本题两份输入独立重算真值，评估 Excel 和面向管理层的 PPTX。"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import tempfile
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


TASK_ID = "cowork-evo/cowork-bench-fy2026-leadership-budget-review"
PDF_NAME = "department_budgets_fy2026.pdf"
EMPLOYEES_NAME = "HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql"
XLSX_NAME = "budget_vs_actuals.xlsx"
PPTX_NAME = "budget_vs_actuals.pptx"

DETAIL_HEADERS = [
    "Department",
    "Approved_Budget",
    "Actual_Expenditure",
    "Variance",
    "Variance_Pct",
    "Headcount",
    "Avg_Salary",
]
TOTAL_LABELS = [
    "Total_Departments",
    "Total_Budget",
    "Total_Actuals",
    "Total_Variance",
    "Most_Under_Budget",
    "Most_Over_Budget",
]

CRITERIA = {
    "R001": (5, True, "两个指定交付文件存在、非空且可解析"),
    "R002": (6, True, "Excel 两张表、表头、部门集与俄文字母序正确"),
    "R003": (12, True, "Excel 的预算、实际、偏差、百分比、人数和平均薪资与输入重算值一致"),
    "R004": (7, True, "PPTX 按标题页、概览页、7 个部门页、Summary 页的顺序组织"),
    "R005": (13, True, "7 个部门页包含全部要求指标，且数值从输入重算相符"),
    "R006": (8, True, "Overview 和 Summary 中的总计、极值部门和状态正确"),
    "R007": (5, True, "Excel 与 PPTX 的跨文件数字和排名一致"),
    "R008": (5, True, "PPTX 为 16:9，使用可编辑原生对象，不是整页图片"),
    "R009": (4, False, "PPTX 能由 LibreOffice 渲染成 10 页 PDF，且按页核心文字可读"),
    "R010": (4, True, "PPTX 不含输入无法支持的额外数值或预测、招聘等断言"),
    "R011": (3, False, "Excel 具有管理层可用的表头、格式、冻结、筛选和列宽"),
    "R012": (3, False, "PPTX 标题层级可读，文字不越界，原生对象数量充足"),
}


def q2(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def command(args: list[str], timeout: int = 240) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, capture_output=True, text=True, timeout=timeout, check=False)


def read_budgets(pdf_path: Path) -> dict[str, Decimal]:
    executable = shutil.which("pdftotext")
    if executable is None:
        raise RuntimeError("缺少 pdftotext")
    completed = command([executable, "-layout", str(pdf_path), "-"])
    if completed.returncode != 0:
        raise RuntimeError((completed.stderr or "pdftotext 失败")[-1000:])
    budgets: dict[str, Decimal] = {}
    pattern = re.compile(r"^(?P<department>\S(?:.*?\S)?)\s+\$(?P<amount>[0-9][0-9,]*\.\d{2})$")
    for raw_line in completed.stdout.splitlines():
        match = pattern.match(raw_line.strip())
        if not match or match.group("department") == "Итого":
            continue
        budgets[match.group("department")] = Decimal(match.group("amount").replace(",", ""))
    if len(budgets) != 7:
        raise ValueError(f"从 PDF 解析到的部门数={len(budgets)}")
    return budgets


def read_employee_aggregate(copy_path: Path) -> dict[str, tuple[int, Decimal]]:
    with copy_path.open(encoding="utf-8", newline="") as stream:
        header = stream.readline().rstrip("\r\n")
        if not header.startswith('COPY sf_data."HR_ANALYTICS__PUBLIC__EMPLOYEES" '):
            raise ValueError("COPY 声明不匹配上游表")
        columns = re.findall(r'"([^"]+)"', header)[1:]
        if len(columns) != 18 or "DEPARTMENT" not in columns or "SALARY" not in columns:
            raise ValueError(f"COPY 字段异常：{columns}")
        department_index = columns.index("DEPARTMENT")
        salary_index = columns.index("SALARY")
        aggregate: dict[str, list[Any]] = defaultdict(lambda: [0, Decimal("0")])
        terminated = False
        for line_number, raw_line in enumerate(stream, start=1):
            line = raw_line.rstrip("\r\n")
            if line == "\\.":
                terminated = True
                break
            fields = line.split("\t")
            if len(fields) != len(columns):
                raise ValueError(f"COPY 第 {line_number} 行字段数异常")
            department = fields[department_index]
            aggregate[department][0] += 1
            aggregate[department][1] += Decimal(fields[salary_index])
        if not terminated:
            raise ValueError("COPY 快照没有终止行")
    return {department: (int(values[0]), q2(Decimal(values[1]))) for department, values in aggregate.items()}


def load_truth(input_dir: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    budgets = read_budgets(input_dir / PDF_NAME)
    actuals = read_employee_aggregate(input_dir / EMPLOYEES_NAME)
    if set(budgets) != set(actuals):
        raise ValueError("两份输入的部门集不一致")
    rows: list[dict[str, Any]] = []
    for department in sorted(budgets):
        headcount, actual = actuals[department]
        budget = q2(budgets[department])
        variance = q2(budget - actual)
        rows.append(
            {
                "department": department,
                "budget": budget,
                "actual": actual,
                "variance": variance,
                "variance_pct": q2(variance / budget * 100),
                "headcount": headcount,
                "avg_salary": q2(actual / headcount),
                "status": "under budget" if variance >= 0 else "over budget",
            }
        )
    total_budget = q2(sum((row["budget"] for row in rows), Decimal("0")))
    total_actual = q2(sum((row["actual"] for row in rows), Decimal("0")))
    most_under = max(rows, key=lambda row: row["variance"])
    most_over = min(rows, key=lambda row: row["variance"])
    totals = {
        "total_departments": len(rows),
        "total_budget": total_budget,
        "total_actual": total_actual,
        "total_variance": q2(total_budget - total_actual),
        "most_under": most_under["department"],
        "most_under_variance": most_under["variance"],
        "most_over": most_over["department"],
        "most_over_variance": most_over["variance"],
        "over_count": sum(1 for row in rows if row["variance"] < 0),
    }
    return rows, totals


def actual_number(raw_cell, cached_cell) -> Decimal | None:
    for value in (cached_cell.value, raw_cell.value):
        if value is None or isinstance(value, bool):
            continue
        if isinstance(value, (int, float, Decimal)):
            return Decimal(str(value))
    return None


def recalc_copy(path: Path, temp_root: Path) -> tuple[Path | None, str]:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if executable is None:
        return None, "缺少 LibreOffice"
    output_dir = temp_root / "xlsx"
    output_dir.mkdir(parents=True, exist_ok=True)
    profile = temp_root / "xlsx-profile"
    completed = command(
        [
            executable,
            "--headless",
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(path),
        ]
    )
    result = output_dir / path.name
    if completed.returncode != 0 or not result.is_file():
        return None, (completed.stderr or completed.stdout or "XLSX 转换失败")[-1000:]
    return result, "LibreOffice 重算及另存成功"


def check_workbook(path: Path, truth_rows: list[dict[str, Any]], totals: dict[str, Any], temp_root: Path) -> tuple[dict[str, tuple[bool, str]], dict[str, dict[str, Decimal]], dict[str, Any]]:
    checks: dict[str, tuple[bool, str]] = {}
    observed: dict[str, dict[str, Decimal]] = {}
    observed_totals: dict[str, Any] = {}
    try:
        raw = load_workbook(path, data_only=False)
    except Exception as exc:
        return {"structure": (False, f"XLSX 无法解析：{exc}"), "values": (False, "未检查")}, observed, observed_totals
    recalculated_path, recalc_note = recalc_copy(path, temp_root)
    try:
        cached = load_workbook(recalculated_path or path, data_only=True)
    except Exception as exc:
        raw.close()
        return {"structure": (False, f"XLSX 缓存值无法解析：{exc}"), "values": (False, "未检查")}, observed, observed_totals

    expected_sheets = ["Department Details", "Totals"]
    structure_errors: list[str] = []
    value_errors: list[str] = []
    if raw.sheetnames != expected_sheets:
        structure_errors.append(f"工作表={raw.sheetnames}")
    if all(name in raw.sheetnames for name in expected_sheets):
        detail_raw = raw["Department Details"]
        detail_cached = cached["Department Details"]
        headers = [str(cell.value or "").strip() for cell in detail_raw[1]][:7]
        if headers != DETAIL_HEADERS:
            structure_errors.append(f"明细表头={headers}")
        populated = [row for row in detail_raw.iter_rows(min_row=2, values_only=False) if row[0].value not in (None, "")]
        if len(populated) != 7:
            structure_errors.append(f"明细行数={len(populated)}")
        names = [str(row[0].value).strip() for row in populated]
        expected_names = [row["department"] for row in truth_rows]
        if names != expected_names:
            structure_errors.append(f"部门顺序={names}，期望={expected_names}")
        truth_by_name = {row["department"]: row for row in truth_rows}
        for row_index, raw_row in enumerate(populated, start=2):
            name = str(raw_row[0].value).strip()
            truth = truth_by_name.get(name)
            if truth is None:
                value_errors.append(f"未知部门 {name}")
                continue
            cached_row = detail_cached[row_index]
            values = [actual_number(raw_row[index], cached_row[index]) for index in range(1, 7)]
            labels = ["budget", "actual", "variance", "variance_pct", "headcount", "avg_salary"]
            tolerances = [Decimal("0.01"), Decimal("0.01"), Decimal("0.01"), Decimal("0.01"), Decimal("0"), Decimal("0.01")]
            found: dict[str, Decimal] = {}
            for label, value, tolerance in zip(labels, values, tolerances):
                expected = Decimal(str(truth[label]))
                if value is None or abs(value - expected) > tolerance:
                    value_errors.append(f"{name} {label}={value}，期望={expected}")
                elif value is not None:
                    found[label] = value
            if found:
                observed[name] = found

        total_raw = raw["Totals"]
        total_cached = cached["Totals"]
        start_row = 2 if normalize(total_raw["A1"].value) == "metric" else 1
        pairs: dict[str, Any] = {}
        for row_index in range(start_row, total_raw.max_row + 1):
            key = str(total_raw.cell(row_index, 1).value or "").strip()
            if not key:
                continue
            raw_value = total_raw.cell(row_index, 2)
            cached_value = total_cached.cell(row_index, 2)
            number = actual_number(raw_value, cached_value)
            pairs[key] = number if number is not None else cached_value.value
        if list(pairs) != TOTAL_LABELS:
            structure_errors.append(f"Totals 标签={list(pairs)}")
        expected_totals = {
            "Total_Departments": Decimal(totals["total_departments"]),
            "Total_Budget": totals["total_budget"],
            "Total_Actuals": totals["total_actual"],
            "Total_Variance": totals["total_variance"],
        }
        for key, expected in expected_totals.items():
            value = pairs.get(key)
            if not isinstance(value, Decimal) or abs(value - Decimal(expected)) > Decimal("0.01"):
                value_errors.append(f"Totals {key}={value}，期望={expected}")
        for key, expected in (
            ("Most_Under_Budget", totals["most_under"]),
            ("Most_Over_Budget", totals["most_over"]),
        ):
            if str(pairs.get(key, "")).strip() != expected:
                value_errors.append(f"Totals {key}={pairs.get(key)!r}，期望={expected!r}")
        observed_totals = pairs

        style_parts = []
        for sheet_name in expected_sheets:
            ws = raw[sheet_name]
            header_row = 1
            header_cells = list(ws[header_row])
            header_ok = all(bool(cell.font.bold) and cell.fill.fill_type == "solid" for cell in header_cells if cell.value is not None)
            freeze_ok = str(ws.freeze_panes) == "A2"
            filter_ok = bool(ws.auto_filter.ref) or bool(ws.tables)
            width_ok = (ws.column_dimensions["A"].width or 0) >= 18
            style_parts.append((sheet_name, header_ok, freeze_ok, filter_ok, width_ok))
        currency_ok = all("0.00" in str(detail_raw[f"B{idx}"].number_format) for idx in range(2, 9))
        style_ok = all(all(part[1:]) for part in style_parts) and currency_ok and recalculated_path is not None
        checks["style"] = (style_ok, f"{style_parts}；currency={currency_ok}；{recalc_note}")
    else:
        checks["style"] = (False, "缺少必需工作表")

    raw.close()
    cached.close()
    checks["structure"] = (not structure_errors, "结构正确" if not structure_errors else "；".join(structure_errors))
    checks["values"] = (not value_errors, "所有 Excel 数值从输入重算相符" if not value_errors else "；".join(value_errors[:16]))
    return checks, observed, observed_totals


def slide_text(slide) -> str:
    parts: list[str] = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False) and shape.text.strip():
            parts.append(shape.text)
        if getattr(shape, "has_table", False):
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        parts.append(cell.text)
    return "\n".join(parts)


def top_text(slide, slide_height: int) -> str:
    return "\n".join(
        shape.text
        for shape in slide.shapes
        if getattr(shape, "has_text_frame", False)
        and shape.text.strip()
        and int(shape.top) <= int(slide_height * 0.24)
    )


def number_tokens(text: str) -> list[Decimal]:
    normalized = str(text).replace("\xa0", " ").replace("\u202f", " ").replace("\u2009", " ")
    pattern = re.compile(r"(?<![\w])[-+]?\$?\d[\d ,.]*\d%?|(?<![\w])[-+]?\$?\d%?")
    values: list[Decimal] = []
    for match in pattern.finditer(normalized):
        token = match.group(0).strip().replace("$", "").replace("%", "").replace(" ", "")
        if not token:
            continue
        if "," in token and "." in token:
            if token.rfind(".") > token.rfind(","):
                token = token.replace(",", "")
            else:
                token = token.replace(".", "").replace(",", ".")
        elif token.count(",") == 1:
            tail = token.split(",", 1)[1]
            token = token.replace(",", "." if len(tail) in {1, 2} else "")
        elif token.count(",") > 1:
            token = token.replace(",", "")
        try:
            values.append(Decimal(token))
        except Exception:
            continue
    return values


def contains_number(text: str, expected: Decimal | int, tolerance: Decimal = Decimal("0.01")) -> bool:
    target = Decimal(str(expected))
    return any(abs(value - target) <= tolerance for value in number_tokens(text))


def check_presentation(path: Path, truth_rows: list[dict[str, Any]], totals: dict[str, Any]) -> tuple[dict[str, tuple[bool, str]], Any, dict[str, str]]:
    checks: dict[str, tuple[bool, str]] = {}
    slide_texts: dict[str, str] = {}
    try:
        prs = Presentation(path)
    except Exception as exc:
        return {"structure": (False, f"PPTX 无法解析：{exc}")}, None, slide_texts
    slides = list(prs.slides)
    structure_errors: list[str] = []
    metric_errors: list[str] = []
    summary_errors: list[str] = []
    if len(slides) != 10:
        structure_errors.append(f"页数={len(slides)}")
    if slides:
        first = normalize(slide_text(slides[0]))
        if "fy2026 budget vs actuals" not in first:
            structure_errors.append("标题页缺少指定标题")
        if "quarterly leadership review - march 2026" not in first:
            structure_errors.append("标题页缺少指定副标题")
    if len(slides) >= 2 and "department overview" not in normalize(top_text(slides[1], int(prs.slide_height))):
        structure_errors.append("第 2 页不是 Department Overview")
    if len(slides) >= 10 and "summary" not in normalize(top_text(slides[9], int(prs.slide_height))):
        structure_errors.append("最后一页不是 Summary")

    for offset, truth in enumerate(truth_rows, start=2):
        if len(slides) <= offset:
            metric_errors.append(f"缺少 {truth['department']} 页")
            continue
        text = slide_text(slides[offset])
        slide_texts[truth["department"]] = text
        if truth["department"] not in top_text(slides[offset], int(prs.slide_height)):
            structure_errors.append(f"第 {offset + 1} 页顶部标题不是 {truth['department']}")
        required_labels = [
            "approved budget",
            "actual expenditure",
            "variance",
            "variance pct",
            "headcount",
            "avg salary",
            truth["status"],
        ]
        missing_labels = [label for label in required_labels if label not in normalize(text)]
        if missing_labels:
            metric_errors.append(f"{truth['department']} 缺少标签={missing_labels}")
        for label, expected, tolerance in (
            ("budget", truth["budget"], Decimal("0.01")),
            ("actual", truth["actual"], Decimal("0.01")),
            ("variance", truth["variance"], Decimal("0.01")),
            ("variance_pct", truth["variance_pct"], Decimal("0.01")),
            ("headcount", truth["headcount"], Decimal("0")),
            ("avg_salary", truth["avg_salary"], Decimal("0.01")),
        ):
            if not contains_number(text, expected, tolerance):
                metric_errors.append(f"{truth['department']} 缺少 {label}={expected}")

    if len(slides) >= 2:
        overview = slide_text(slides[1])
        slide_texts["overview"] = overview
        for label, expected, tolerance in (
            ("departments", totals["total_departments"], Decimal("0")),
            ("total budget", totals["total_budget"], Decimal("0.01")),
            ("total actual", totals["total_actual"], Decimal("0.01")),
            ("total variance", totals["total_variance"], Decimal("0.01")),
        ):
            if not contains_number(overview, expected, tolerance):
                summary_errors.append(f"Overview 缺少 {label}={expected}")
    if len(slides) >= 10:
        summary = slide_text(slides[9])
        slide_texts["summary"] = summary
        if totals["most_under"] not in summary or totals["most_over"] not in summary:
            summary_errors.append("汇总页极值部门不正确")
        for label, expected in (
            ("most-under variance", totals["most_under_variance"]),
            ("most-over variance", totals["most_over_variance"]),
        ):
            if not contains_number(summary, expected):
                summary_errors.append(f"Summary 缺少 {label}={expected}")
        if totals["over_count"] == 0 and "all 7 departments are under budget" not in normalize(summary):
            summary_errors.append("汇总页没有正确表达全部部门未超支")
        if totals["over_count"] > 0:
            expected_phrase = f"{totals['over_count']} of 7 department"
            if expected_phrase not in normalize(summary):
                summary_errors.append("汇总页超支部门数不正确")

    slide_w, slide_h = int(prs.slide_width), int(prs.slide_height)
    ratio_ok = abs(slide_w / slide_h - 16 / 9) <= 0.01
    native_count = 0
    text_count = 0
    large_pictures: list[str] = []
    out_of_bounds: list[str] = []
    title_size_errors: list[str] = []
    for page_number, slide in enumerate(slides, start=1):
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if int(shape.width) * int(shape.height) >= slide_w * slide_h * 0.60:
                    large_pictures.append(f"p{page_number}:{shape.name}")
            else:
                native_count += 1
            if getattr(shape, "has_text_frame", False) and shape.text.strip():
                text_count += 1
                tolerance = 36000
                if (
                    int(shape.left) < -tolerance
                    or int(shape.top) < -tolerance
                    or int(shape.left + shape.width) > slide_w + tolerance
                    or int(shape.top + shape.height) > slide_h + tolerance
                ):
                    out_of_bounds.append(f"p{page_number}:{shape.name}")
        if slide.shapes:
            title_candidates = [
                shape
                for shape in slide.shapes
                if getattr(shape, "has_text_frame", False)
                and shape.text.strip()
                and int(shape.top) <= slide_h * 0.24
            ]
            if title_candidates:
                sizes = [
                    run.font.size.pt
                    for shape in title_candidates
                    for paragraph in shape.text_frame.paragraphs
                    for run in paragraph.runs
                    if run.text.strip() and run.font.size is not None
                ]
                if sizes and max(sizes) < 24:
                    title_size_errors.append(f"p{page_number}:max={max(sizes):.1f}pt")
    editability_ok = ratio_ok and native_count >= 70 and text_count >= 35 and not large_pictures
    presentation_ok = not out_of_bounds and not title_size_errors
    checks["structure"] = (not structure_errors, "页面顺序和标题正确" if not structure_errors else "；".join(structure_errors))
    checks["metrics"] = (not metric_errors, "7 个部门页的 6 项指标及状态均正确" if not metric_errors else "；".join(metric_errors[:20]))
    checks["summary"] = (not summary_errors, "Overview 与 Summary 均与输入重算真值一致" if not summary_errors else "；".join(summary_errors))
    checks["editability"] = (editability_ok, f"16:9={ratio_ok}；原生对象={native_count}；文字对象={text_count}；整页大图={large_pictures}")
    checks["presentation"] = (presentation_ok, f"越界={out_of_bounds}；标题字号问题={title_size_errors}")
    return checks, prs, slide_texts


def grounding_check(prs, truth_rows: list[dict[str, Any]], totals: dict[str, Any]) -> tuple[bool, str]:
    texts = [slide_text(slide) for slide in prs.slides]
    unsupported_phrases = re.compile(
        r"\b(next quarter|forecast|guarantee(?:d)?|will hire|hiring plan|layoff|headcount reduction|"
        r"recommend(?:ation)?|market growth|caused by|target savings)\b",
        flags=re.I,
    )
    phrase_hits = [f"p{index + 1}:{match.group(0)}" for index, text in enumerate(texts) if (match := unsupported_phrases.search(text))]

    allowed_by_page: list[list[Decimal]] = [
        [Decimal("2026")],
        [
            Decimal("2026"),
            Decimal(totals["total_departments"]),
            totals["total_budget"],
            totals["total_actual"],
            totals["total_variance"],
        ],
    ]
    for row in truth_rows:
        allowed_by_page.append(
            [
                Decimal("2026"),
                row["budget"],
                row["actual"],
                row["variance"],
                row["variance_pct"],
                Decimal(row["headcount"]),
                row["avg_salary"],
            ]
        )
    most_under = next(row for row in truth_rows if row["department"] == totals["most_under"])
    most_over = next(row for row in truth_rows if row["department"] == totals["most_over"])
    allowed_by_page.append(
        [
            Decimal("7"),
            Decimal(totals["over_count"]),
            most_under["budget"],
            most_under["actual"],
            most_under["variance"],
            most_over["budget"],
            most_over["actual"],
            most_over["variance"],
        ]
    )
    numeric_hits: list[str] = []
    for page_index, text in enumerate(texts):
        allowed = allowed_by_page[page_index] if page_index < len(allowed_by_page) else []
        for value in number_tokens(text):
            if not any(abs(value - expected) <= Decimal("0.01") for expected in allowed):
                numeric_hits.append(f"p{page_index + 1}:{value}")
    ok = not phrase_hits and not numeric_hits
    return ok, f"无来源短语={phrase_hits}；额外数值={numeric_hits[:20]}"


def cross_file_check(observed: dict[str, dict[str, Decimal]], observed_totals: dict[str, Any], slide_texts: dict[str, str], truth_rows: list[dict[str, Any]], totals: dict[str, Any]) -> tuple[bool, str]:
    errors: list[str] = []
    for row in truth_rows:
        department = row["department"]
        excel_row = observed.get(department)
        text = slide_texts.get(department, "")
        if not excel_row or not text:
            errors.append(f"{department} 无法跨文件比较")
            continue
        for field in ("budget", "actual", "variance", "variance_pct", "headcount", "avg_salary"):
            value = excel_row.get(field)
            if value is None or not contains_number(text, value, Decimal("0.01") if field != "headcount" else Decimal("0")):
                errors.append(f"{department} {field} 在 Excel/PPTX 不一致")
    if str(observed_totals.get("Most_Under_Budget", "")) != totals["most_under"]:
        errors.append("Excel 最大正偏差部门与真值不一致")
    if str(observed_totals.get("Most_Over_Budget", "")) != totals["most_over"]:
        errors.append("Excel 最小偏差部门与真值不一致")
    summary = slide_texts.get("summary", "")
    if totals["most_under"] not in summary or totals["most_over"] not in summary:
        errors.append("PPTX Summary 排名与 Excel 不一致")
    return not errors, "Excel 与 PPTX 逐部门交叉一致" if not errors else "；".join(errors[:16])


def render_check(path: Path, truth_rows: list[dict[str, Any]], temp_root: Path) -> tuple[bool, str]:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    pdfinfo = shutil.which("pdfinfo")
    pdftotext = shutil.which("pdftotext")
    if not executable or not pdfinfo or not pdftotext:
        return False, "缺少 LibreOffice/pdfinfo/pdftotext"
    output_dir = temp_root / "ppt-render"
    output_dir.mkdir(parents=True, exist_ok=True)
    profile = temp_root / "ppt-profile"
    completed = command(
        [
            executable,
            "--headless",
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_dir),
            str(path),
        ]
    )
    pdf_path = output_dir / f"{path.stem}.pdf"
    if completed.returncode != 0 or not pdf_path.is_file():
        return False, (completed.stderr or completed.stdout or "PPTX 渲染失败")[-1000:]
    info = command([pdfinfo, str(pdf_path)], timeout=60)
    text = command([pdftotext, "-layout", str(pdf_path), "-"], timeout=120)
    page_match = re.search(r"^Pages:\s+(\d+)", info.stdout, flags=re.M)
    pages = [page for page in text.stdout.split("\f") if page.strip()]
    errors: list[str] = []
    if not page_match or int(page_match.group(1)) != 10 or len(pages) != 10:
        errors.append(f"PDF 页数={page_match.group(1) if page_match else '?'}，文本页={len(pages)}")
    if pages and "fy2026 budget vs actuals" not in normalize(pages[0]):
        errors.append("渲染首页缺少标题")
    for index, row in enumerate(truth_rows, start=2):
        if index >= len(pages) or row["department"] not in pages[index]:
            errors.append(f"渲染第 {index + 1} 页缺少 {row['department']}")
    if len(pages) >= 10 and "summary" not in normalize(pages[9]):
        errors.append("渲染末页缺少 Summary")
    return not errors, "LibreOffice 成功渲染 10 页，按页标题可读" if not errors else "；".join(errors)


def result_item(criterion_id: str, passed: bool, evidence: str) -> dict[str, Any]:
    weight, hurdle, description = CRITERIA[criterion_id]
    return {
        "id": criterion_id,
        "passed": bool(passed),
        "weight": weight,
        "hurdle": hurdle,
        "score": 1.0 if passed else 0.0,
        "description": description,
        "evidence": evidence,
    }


def evaluate(artifact_dir: Path, input_dir: Path) -> dict[str, Any]:
    xlsx_path = artifact_dir / XLSX_NAME
    pptx_path = artifact_dir / PPTX_NAME
    file_errors: list[str] = []
    for path, minimum in ((xlsx_path, 3000), (pptx_path, 10000)):
        if not path.is_file():
            file_errors.append(f"缺少 {path.name}")
        elif path.stat().st_size < minimum:
            file_errors.append(f"{path.name} 过小：{path.stat().st_size}")

    items: list[dict[str, Any]] = []
    if file_errors:
        items.append(result_item("R001", False, "；".join(file_errors)))
        for criterion_id in CRITERIA:
            if criterion_id != "R001":
                items.append(result_item(criterion_id, False, "必需交付物缺失，未评估"))
        return finalize(items, artifact_dir, input_dir)

    try:
        truth_rows, totals = load_truth(input_dir)
    except Exception as exc:
        items.append(result_item("R001", False, f"评测输入无法重算：{exc}"))
        for criterion_id in CRITERIA:
            if criterion_id != "R001":
                items.append(result_item(criterion_id, False, "评测输入异常"))
        return finalize(items, artifact_dir, input_dir)

    with tempfile.TemporaryDirectory(prefix="cowork-evo-budget-judge-") as temp_dir:
        temp_root = Path(temp_dir)
        workbook_checks, observed, observed_totals = check_workbook(xlsx_path, truth_rows, totals, temp_root)
        presentation_checks, prs, slide_texts = check_presentation(pptx_path, truth_rows, totals)
        parse_ok = prs is not None and "structure" in workbook_checks
        items.append(result_item("R001", parse_ok, "XLSX 与 PPTX 均可解析" if parse_ok else "至少一个文件无法解析"))
        items.append(result_item("R002", *workbook_checks.get("structure", (False, "未检查"))))
        items.append(result_item("R003", *workbook_checks.get("values", (False, "未检查"))))
        items.append(result_item("R004", *presentation_checks.get("structure", (False, "未检查"))))
        items.append(result_item("R005", *presentation_checks.get("metrics", (False, "未检查"))))
        items.append(result_item("R006", *presentation_checks.get("summary", (False, "未检查"))))
        cross = cross_file_check(observed, observed_totals, slide_texts, truth_rows, totals)
        items.append(result_item("R007", *cross))
        items.append(result_item("R008", *presentation_checks.get("editability", (False, "未检查"))))
        render = render_check(pptx_path, truth_rows, temp_root)
        items.append(result_item("R009", *render))
        grounding = grounding_check(prs, truth_rows, totals) if prs is not None else (False, "PPTX 无法解析")
        items.append(result_item("R010", *grounding))
        items.append(result_item("R011", *workbook_checks.get("style", (False, "未检查"))))
        items.append(result_item("R012", *presentation_checks.get("presentation", (False, "未检查"))))

    return finalize(items, artifact_dir, input_dir)


def finalize(items: list[dict[str, Any]], artifact_dir: Path, input_dir: Path) -> dict[str, Any]:
    total_weight = sum(item["weight"] for item in items)
    earned = sum(item["weight"] * item["score"] for item in items)
    raw_score = earned / total_weight if total_weight else 0.0
    hurdle_failures = [item["id"] for item in items if item["hurdle"] and not item["passed"]]
    score = min(raw_score, 0.49) if hurdle_failures else raw_score
    score = round(max(0.0, min(1.0, score)), 6)
    return {
        "task_id": TASK_ID,
        "score": score,
        "reward": score,
        "passed": score >= 0.85 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
        "artifact_dir": str(artifact_dir),
        "input_dir": str(input_dir),
        "criteria": items,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--data-dir", type=Path, default=Path("/app/data/input_files"))
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir, args.data_dir)
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(payload, encoding="utf-8")
    print(payload, end="")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
