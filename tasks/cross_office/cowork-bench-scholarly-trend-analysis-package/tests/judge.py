#!/usr/bin/env python3
"""对论文 JSON、HTML 面板、处理脚本、Excel 与 PPT 做端到端确定性评分。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from decimal import Decimal, ROUND_HALF_UP
from html.parser import HTMLParser
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


PROCESSOR_NAME = "scholarly_trend_processor.py"
JSON_NAME = "scholarly_trend_results.json"
XLSX_NAME = "Trend_Analysis_Report.xlsx"
PPTX_NAME = "Trend_Analysis_Presentation.pptx"
OUTPUT_NAMES = [PROCESSOR_NAME, JSON_NAME, XLSX_NAME, PPTX_NAME]
TASK_ID = "cowork-evo/cowork-bench-scholarly-trend-analysis-package"
RUBRIC_VERSION = "cowork-bench-scholarly-trend-analysis-package-v0.1"
JUDGE_VERSION = "cowork-bench-scholarly-trend-analysis-package-v0.1"
PASS_THRESHOLD = 0.8
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.35,
    "R003": 0.45,
    "R004": 0.35,
    "R005": 0.5,
    "R006": 0.35,
    "R007": 0.6,
    "R009": 0.55,
    "R011": 0.65,
}
DATA_HEADERS = ["Paper_ID", "Title", "Area", "Citations", "Relevance_Score"]
METRIC_ORDER = ["Total_Papers", "Avg_Citations", "Top_Area", "Avg_Relevance"]
RECOMMENDATION_HEADERS = ["Priority", "Action", "Area"]

CRITERION_SPECS = {
    "R001": ("脚本、JSON、Excel 和 PPT 四个交付物完整可读", 5, "hurdle"),
    "R002": ("处理脚本在原始输入上按规定 CLI 回放成功并输出真值", 7, "hurdle"),
    "R003": ("处理脚本在变异 HTML 上产生随输入改变的正确结果", 7, "hurdle"),
    "R004": ("交付 JSON 的4篇论文与匹配指标精确正确", 8, "hurdle"),
    "R005": ("交付 JSON 的汇总指标和建议正确", 6, "hurdle"),
    "R006": ("Excel 三张表结构和 Data_Analysis 真值正确", 8, "hurdle"),
    "R007": ("Excel Metrics 使用可审计公式且重算值正确", 7, "hurdle"),
    "R008": ("Excel 与 JSON 的建议完全一致", 4, "positive"),
    "R009": ("PPT 的页面、方法、论文、指标和建议内容完整正确", 7, "hurdle"),
    "R010": ("PPT 为原生可编辑内容，无整页大图、越界或空白页", 4, "positive"),
    "R011": ("JSON、Excel 和 PPT 的论文、领域与指标跨文件一致", 6, "hurdle"),
    "R012": ("Excel 和 PPT 可由 LibreOffice 重算、另存并渲染", 3, "positive"),
    "R013": ("Excel 表格样式、冻结、筛选、列宽与数字格式完整", 3, "positive"),
    "R014": ("任一来源匹配值或汇总指标与输入真值冲突", -14, "penalty"),
    "R015": ("处理脚本不可回放或不随变异 HTML 改变", -12, "penalty"),
    "R016": ("交付物缺失、损坏、截图化或跨文件冲突", -12, "penalty"),
    "R017": ("交付中编造了输入不支持的 gap、趋势方向或差值", -18, "penalty"),
}


def criterion(identifier: str, passed: bool, evidence: str = "") -> dict:
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


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_cell = False
        self.parts: list[str] = []
        self.row: list[str] = []
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() in ("td", "th"):
            self.in_cell = True
            self.parts = []

    def handle_data(self, data: str) -> None:
        if self.in_cell:
            self.parts.append(data)

    def handle_endtag(self, tag: str) -> None:
        lowered = tag.lower()
        if lowered in ("td", "th") and self.in_cell:
            self.row.append(" ".join("".join(self.parts).split()))
            self.in_cell = False
        elif lowered == "tr" and self.row:
            self.rows.append(self.row)
            self.row = []


def parse_dashboard(path: Path) -> dict[str, dict]:
    parser = TableParser()
    parser.feed(path.read_text(encoding="utf-8"))
    headers = ["Research_Area", "Publication_Count_2025", "Avg_Citations", "Trending_Score"]
    for index, row in enumerate(parser.rows):
        if row == headers:
            output = {}
            for values in parser.rows[index + 1 :]:
                if len(values) != 4:
                    break
                item = dict(zip(headers, values))
                output[item["Research_Area"]] = {
                    "Avg_Citations": float(item["Avg_Citations"]),
                    "Trending_Score": float(item["Trending_Score"]),
                }
            if output:
                return output
    raise ValueError("未找到面板表格")


def round_integer(value: float) -> int:
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def derive_expected(papers_path: Path, dashboard_path: Path) -> dict:
    papers = json.loads(papers_path.read_text(encoding="utf-8"))
    dashboard = parse_dashboard(dashboard_path)
    rows = []
    for paper in papers:
        area = str(paper["Area"])
        benchmark = dashboard[area]
        rows.append(
            {
                "Paper_ID": str(paper["Paper_ID"]),
                "Title": str(paper["Title"]),
                "Area": area,
                "Citations": benchmark["Avg_Citations"],
                "Relevance_Score": benchmark["Trending_Score"],
            }
        )
    rows.sort(key=lambda item: item["Area"])
    top = max(rows, key=lambda item: item["Relevance_Score"])
    return {
        "data_analysis": rows,
        "metrics": {
            "Total_Papers": len(rows),
            "Avg_Citations": round_integer(sum(item["Citations"] for item in rows) / len(rows)),
            "Top_Area": top["Area"],
            "Avg_Relevance": round_integer(sum(item["Relevance_Score"] for item in rows) / len(rows)),
        },
    }


def normalized(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def close(actual, expected, tolerance: float = 0.001) -> bool:
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return False


def validate_rows(actual, expected) -> tuple[bool, list]:
    errors = []
    if not isinstance(actual, list) or len(actual) != len(expected):
        return False, [("row_count", len(actual) if isinstance(actual, list) else type(actual).__name__, len(expected))]
    for index, (actual_row, expected_row) in enumerate(zip(actual, expected)):
        if not isinstance(actual_row, dict) or list(actual_row) != DATA_HEADERS:
            errors.append((index, "keys", list(actual_row) if isinstance(actual_row, dict) else type(actual_row).__name__, DATA_HEADERS))
            continue
        for key in ("Paper_ID", "Title", "Area"):
            if str(actual_row.get(key) or "").strip() != expected_row[key]:
                errors.append((index, key, actual_row.get(key), expected_row[key]))
        for key in ("Citations", "Relevance_Score"):
            if not close(actual_row.get(key), expected_row[key]):
                errors.append((index, key, actual_row.get(key), expected_row[key]))
    return not errors, errors


def validate_metrics(actual, expected) -> tuple[bool, list]:
    errors = []
    if not isinstance(actual, dict) or list(actual) != METRIC_ORDER:
        errors.append(("keys", list(actual) if isinstance(actual, dict) else type(actual).__name__, METRIC_ORDER))
        return False, errors
    for key in METRIC_ORDER:
        if isinstance(expected[key], str):
            if str(actual.get(key) or "").strip() != expected[key]:
                errors.append((key, actual.get(key), expected[key]))
        elif not close(actual.get(key), expected[key], 0):
            errors.append((key, actual.get(key), expected[key]))
    return not errors, errors


def validate_recommendations(actual, areas: set[str]) -> tuple[bool, list]:
    errors = []
    if not isinstance(actual, list) or len(actual) < 2:
        return False, [("row_count", len(actual) if isinstance(actual, list) else type(actual).__name__, ">=2")]
    for index, row in enumerate(actual):
        if not isinstance(row, dict) or list(row) != RECOMMENDATION_HEADERS:
            errors.append((index, "keys", list(row) if isinstance(row, dict) else type(row).__name__))
            continue
        if any(not str(row.get(key) or "").strip() for key in RECOMMENDATION_HEADERS):
            errors.append((index, "blank", row))
        if str(row.get("Area") or "").strip() not in areas:
            errors.append((index, "unknown_area", row.get("Area")))
    return not errors, errors


def load_result(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or list(data) != ["data_analysis", "metrics", "recommendations"]:
        raise ValueError(f"JSON 顶层键错误：{list(data) if isinstance(data, dict) else type(data).__name__}")
    return data


def run_processor(script: Path, papers: Path, dashboard: Path, output: Path) -> tuple[dict | None, str]:
    try:
        completed = subprocess.run(
            [sys.executable, str(script.resolve()), "--papers", str(papers.resolve()), "--dashboard", str(dashboard.resolve()), "--output", str(output.resolve())],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
            cwd=output.parent,
        )
        if completed.returncode != 0 or not output.is_file():
            return None, f"exit={completed.returncode}，stdout={completed.stdout[-500:]}，stderr={completed.stderr[-500:]}"
        return load_result(output), "exit=0，输出 JSON 可读"
    except Exception as error:
        return None, str(error)


def read_sheet_rows(sheet, headers: list[str]) -> tuple[list[dict], list[str]]:
    values = list(sheet.iter_rows(values_only=True))
    if not values:
        return [], headers
    actual_headers = [str(value).strip() if value is not None else "" for value in values[0]]
    missing = [header for header in headers if header not in actual_headers]
    if missing:
        return [], missing
    positions = {header: actual_headers.index(header) for header in headers}
    rows = []
    for value_row in values[1:]:
        if not value_row or all(value is None or str(value).strip() == "" for value in value_row):
            continue
        rows.append({header: value_row[positions[header]] for header in headers})
    return rows, []


def formula_norm(value) -> str:
    return re.sub(r"[\s$']", "", str(value or "")).upper()


def metric_formulas_ok(workbook) -> tuple[bool, dict]:
    sheet = workbook["Metrics"]
    formulas = {f"B{row}": str(sheet[f"B{row}"].value or "") for row in range(2, 6)}
    normalized_formulas = {cell: formula_norm(value) for cell, value in formulas.items()}
    checks = [
        normalized_formulas["B2"].startswith("=") and "DATA_ANALYSIS!A2:A5" in normalized_formulas["B2"] and ("COUNTA" in normalized_formulas["B2"] or "ROWS" in normalized_formulas["B2"]),
        normalized_formulas["B3"].startswith("=") and "AVERAGE" in normalized_formulas["B3"] and "DATA_ANALYSIS!D2:D5" in normalized_formulas["B3"],
        normalized_formulas["B4"].startswith("=") and "DATA_ANALYSIS!C2:C5" in normalized_formulas["B4"] and "MAX" in normalized_formulas["B4"] and "DATA_ANALYSIS!E2:E5" in normalized_formulas["B4"],
        normalized_formulas["B5"].startswith("=") and "AVERAGE" in normalized_formulas["B5"] and "DATA_ANALYSIS!E2:E5" in normalized_formulas["B5"],
    ]
    return all(checks), formulas


def libreoffice_prepare(xlsx: Path, pptx: Path, root: Path) -> tuple[Path | None, dict[str, str], dict[str, int], str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    pdfinfo = shutil.which("pdfinfo")
    if libreoffice is None or pdftotext is None or pdfinfo is None:
        return None, {}, {}, "缺少 libreoffice/pdftotext/pdfinfo"
    source_dir = root / "source"
    recalc_dir = root / "recalc"
    pdf_dir = root / "pdf"
    source_dir.mkdir(); recalc_dir.mkdir(); pdf_dir.mkdir()
    xlsx_copy = source_dir / XLSX_NAME
    pptx_copy = source_dir / PPTX_NAME
    shutil.copy2(xlsx, xlsx_copy); shutil.copy2(pptx, pptx_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()
    recalc = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(xlsx_copy)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    recalculated = recalc_dir / XLSX_NAME
    if recalc.returncode != 0 or not recalculated.is_file():
        return None, {}, {}, (recalc.stderr or recalc.stdout or "Excel 重算失败")[-1000:]
    texts = {}; pages = {}
    for source in (recalculated, pptx_copy):
        converted = subprocess.run(
            [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(source)],
            capture_output=True, text=True, timeout=300, check=False,
        )
        pdf_path = pdf_dir / f"{source.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return recalculated, texts, pages, (converted.stderr or converted.stdout or f"{source.name} 渲染失败")[-1000:]
        extracted = subprocess.run([pdftotext, "-raw", str(pdf_path), "-"], capture_output=True, text=True, timeout=120, check=False)
        info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=120, check=False)
        if extracted.returncode != 0 or info.returncode != 0:
            return recalculated, texts, pages, f"{source.name} PDF 检查失败"
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        texts[source.name] = extracted.stdout
        pages[source.name] = int(match.group(1)) if match else 0
    return recalculated, texts, pages, "Excel 重算和 Excel/PPT PDF 渲染成功"


def ppt_text_and_structure(path: Path) -> tuple[Presentation, list[str], int, int, list, list]:
    presentation = Presentation(path)
    slide_texts = []
    native_objects = 0
    pictures = 0
    oversized_pictures = []
    out_of_bounds = []
    width = presentation.slide_width
    height = presentation.slide_height
    for slide_index, slide in enumerate(presentation.slides, start=1):
        parts = []
        for shape in slide.shapes:
            native_objects += 1
            if getattr(shape, "has_text_frame", False):
                parts.append(shape.text)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures += 1
                if shape.width * shape.height >= width * height * 0.6:
                    oversized_pictures.append(slide_index)
            if shape.left < -1000 or shape.top < -1000 or shape.left + shape.width > width + 1000 or shape.top + shape.height > height + 1000:
                out_of_bounds.append((slide_index, shape.name))
        slide_texts.append("\n".join(parts))
    return presentation, slide_texts, native_objects, pictures, oversized_pictures, out_of_bounds


def unsupported_gap_claims(*texts: str) -> list[str]:
    combined = "\n".join(texts)
    patterns = {
        "numeric_gap": r"(?:gap|разрыв|отклонен)\s*(?:=|:|составляет)?\s*-?\d+(?:[.,]\d+)?\s*%",
        "trend_direction": r"(?:large language models|computer vision|reinforcement learning|federated learning).{0,80}(?:improving|declining|stable|растёт|снижается|стабильн)",
    }
    return [label for label, pattern in patterns.items() if re.search(pattern, combined, re.IGNORECASE | re.DOTALL)]


def evaluate(artifact_dir: Path) -> dict:
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results = []
    if not all(path.is_file() for path in paths.values()):
        results.append(criterion("R001", False, f"文件存在性={{{', '.join(f'{name}: {path.is_file()}' for name, path in paths.items())}}}"))
        return build_result(results, list(paths.values()))

    task_root = Path(__file__).resolve().parents[1]
    input_dir = Path("/app/input_files") if Path("/app/input_files").is_dir() else task_root / "data" / "input_files"
    expected = derive_expected(input_dir / "scholarly_papers.json", input_dir / "dashboard.html")
    try:
        delivered = load_result(paths[JSON_NAME])
        if not zipfile.is_zipfile(paths[XLSX_NAME]) or not zipfile.is_zipfile(paths[PPTX_NAME]):
            raise ValueError("Excel 或 PPT 不是有效 OOXML ZIP")
        script_text = paths[PROCESSOR_NAME].read_text(encoding="utf-8")
        if len(script_text.strip()) < 300:
            raise ValueError("处理脚本内容过短")
    except Exception as error:
        results.append(criterion("R001", False, f"无法读取交付物：{error}"))
        return build_result(results, list(paths.values()))
    results.append(criterion("R001", True, f"四个文件均可读；脚本字符={len(script_text)}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-scholar-script-") as temp_dir:
        temp = Path(temp_dir)
        canonical_result, canonical_evidence = run_processor(
            paths[PROCESSOR_NAME], input_dir / "scholarly_papers.json", input_dir / "dashboard.html", temp / "canonical.json"
        )
        canonical_rows_ok, canonical_row_errors = validate_rows(canonical_result.get("data_analysis") if canonical_result else None, expected["data_analysis"])
        canonical_metrics_ok, canonical_metric_errors = validate_metrics(canonical_result.get("metrics") if canonical_result else None, expected["metrics"])
        canonical_recs_ok, canonical_rec_errors = validate_recommendations(
            canonical_result.get("recommendations") if canonical_result else None,
            {row["Area"] for row in expected["data_analysis"]},
        )
        canonical_ok = canonical_result is not None and canonical_rows_ok and canonical_metrics_ok and canonical_recs_ok

        mutated_html = (input_dir / "dashboard.html").read_text(encoding="utf-8")
        mutated_html = mutated_html.replace("<td>10.8</td>", "<td>14.8</td>", 1).replace("<td>82</td>", "<td>86</td>", 1)
        mutated_path = temp / "dashboard-mutated.html"
        mutated_path.write_text(mutated_html, encoding="utf-8")
        mutated_expected = derive_expected(input_dir / "scholarly_papers.json", mutated_path)
        mutated_result, mutated_evidence = run_processor(
            paths[PROCESSOR_NAME], input_dir / "scholarly_papers.json", mutated_path, temp / "mutated.json"
        )
        mutated_rows_ok, mutated_row_errors = validate_rows(mutated_result.get("data_analysis") if mutated_result else None, mutated_expected["data_analysis"])
        mutated_metrics_ok, mutated_metric_errors = validate_metrics(mutated_result.get("metrics") if mutated_result else None, mutated_expected["metrics"])
        mutation_changed = (
            mutated_result is not None
            and canonical_result is not None
            and mutated_result.get("data_analysis") != canonical_result.get("data_analysis")
            and mutated_result.get("metrics") != canonical_result.get("metrics")
        )
        mutated_ok = mutated_result is not None and mutated_rows_ok and mutated_metrics_ok and mutation_changed

    results.append(criterion("R002", canonical_ok, f"{canonical_evidence}，行错={canonical_row_errors}，指标错={canonical_metric_errors}，建议错={canonical_rec_errors}"))
    results.append(criterion("R003", mutated_ok, f"{mutated_evidence}，行错={mutated_row_errors}，指标错={mutated_metric_errors}，输出已改变={mutation_changed}，变异指标={mutated_expected['metrics']}"))

    json_rows_ok, json_row_errors = validate_rows(delivered["data_analysis"], expected["data_analysis"])
    json_metrics_ok, json_metric_errors = validate_metrics(delivered["metrics"], expected["metrics"])
    json_recs_ok, json_rec_errors = validate_recommendations(delivered["recommendations"], {row["Area"] for row in expected["data_analysis"]})
    results.append(criterion("R004", json_rows_ok, f"错误={json_row_errors}"))
    results.append(criterion("R005", json_metrics_ok and json_recs_ok, f"指标错={json_metric_errors}，建议错={json_rec_errors}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-scholar-office-") as temp_dir:
        recalculated, rendered_texts, rendered_pages, render_evidence = libreoffice_prepare(paths[XLSX_NAME], paths[PPTX_NAME], Path(temp_dir))
        workbook_path = recalculated or paths[XLSX_NAME]
        try:
            value_book = load_workbook(workbook_path, data_only=True)
            formula_book = load_workbook(paths[XLSX_NAME], data_only=False)
            sheet_names = value_book.sheetnames
            data_rows, data_missing = read_sheet_rows(value_book["Data_Analysis"], DATA_HEADERS) if "Data_Analysis" in sheet_names else ([], DATA_HEADERS)
            metric_rows, metric_missing = read_sheet_rows(value_book["Metrics"], ["Metric", "Value"]) if "Metrics" in sheet_names else ([], ["Metric", "Value"])
            recommendation_rows, recommendation_missing = read_sheet_rows(value_book["Recommendations"], RECOMMENDATION_HEADERS) if "Recommendations" in sheet_names else ([], RECOMMENDATION_HEADERS)
            formulas_ok, formula_evidence = metric_formulas_ok(formula_book) if "Metrics" in formula_book.sheetnames else (False, {})
            formatting = {}
            for sheet_name in ("Data_Analysis", "Metrics", "Recommendations"):
                if sheet_name not in formula_book.sheetnames:
                    formatting[sheet_name] = {"exists": False}
                    continue
                sheet = formula_book[sheet_name]
                formatting[sheet_name] = {
                    "tables": len(sheet.tables),
                    "freeze": str(sheet.freeze_panes or ""),
                    "header_bold": all(cell.font.bold for cell in sheet[1]),
                    "header_fill": all(cell.fill.fill_type == "solid" for cell in sheet[1]),
                    "wide_columns": all((sheet.column_dimensions[column].width or 0) >= 14 for column in ("ABCDE" if sheet_name == "Data_Analysis" else "AB" if sheet_name == "Metrics" else "ABC")),
                }
            value_book.close(); formula_book.close()
        except Exception as error:
            results.append(criterion("R006", False, f"Excel 读取失败：{error}"))
            return build_result(results, list(paths.values()))

        try:
            presentation, slide_texts, native_objects, pictures, oversized_pictures, out_of_bounds = ppt_text_and_structure(paths[PPTX_NAME])
            slide_width = presentation.slide_width
            slide_height = presentation.slide_height
            slide_count = len(presentation.slides)
        except Exception as error:
            results.append(criterion("R009", False, f"PPT 读取失败：{error}"))
            return build_result(results, list(paths.values()))

    excel_rows_ok, excel_row_errors = validate_rows(data_rows, expected["data_analysis"])
    excel_structure_ok = sheet_names == ["Data_Analysis", "Metrics", "Recommendations"] and not data_missing and not metric_missing and not recommendation_missing
    results.append(criterion("R006", excel_structure_ok and excel_rows_ok, f"工作表={sheet_names}，缺表头={data_missing + metric_missing + recommendation_missing}，数据错={excel_row_errors}"))

    actual_excel_metrics = {str(row["Metric"] or "").strip(): row["Value"] for row in metric_rows}
    excel_metric_errors = []
    for metric in METRIC_ORDER:
        if metric not in actual_excel_metrics:
            excel_metric_errors.append((metric, "缺失", expected["metrics"][metric]))
        elif isinstance(expected["metrics"][metric], str):
            if str(actual_excel_metrics[metric] or "").strip() != expected["metrics"][metric]:
                excel_metric_errors.append((metric, actual_excel_metrics[metric], expected["metrics"][metric]))
        elif not close(actual_excel_metrics[metric], expected["metrics"][metric], 0):
            excel_metric_errors.append((metric, actual_excel_metrics[metric], expected["metrics"][metric]))
    excel_metrics_ok = list(actual_excel_metrics) == METRIC_ORDER and not excel_metric_errors
    results.append(criterion("R007", formulas_ok and excel_metrics_ok and recalculated is not None, f"公式={formula_evidence}，指标错={excel_metric_errors}，LibreOffice重算={recalculated is not None}"))

    canonical_excel_recs = [
        {header: str(row[header]).strip() for header in RECOMMENDATION_HEADERS}
        for row in recommendation_rows
    ]
    canonical_json_recs = [
        {header: str(row[header]).strip() for header in RECOMMENDATION_HEADERS}
        for row in delivered["recommendations"]
    ]
    recommendations_cross_ok = canonical_excel_recs == canonical_json_recs
    results.append(criterion("R008", recommendations_cross_ok, f"Excel={canonical_excel_recs}，JSON={canonical_json_recs}"))

    ppt_blob = "\n".join(slide_texts)
    ppt_norm = normalized(ppt_blob)
    aspect = slide_width / slide_height if slide_height else 0
    missing_papers = []
    for row in expected["data_analysis"]:
        for token in (row["Paper_ID"], row["Title"], row["Area"], f"{row['Citations']:.1f}", f"{row['Relevance_Score']:.0f}"):
            if normalized(token) not in ppt_norm:
                missing_papers.append(token)
    metric_tokens = ["total_papers", "avg_citations", "top_area", "avg_relevance", "4", "10", "79", "large language models"]
    missing_metrics = [token for token in metric_tokens if token not in ppt_norm]
    method_ok = all(token in ppt_norm for token in ("scholarly_papers.json", "dashboard.html", "area = research_area", "avg_citations", "trending_score"))
    recommendation_areas_missing = [row["Area"] for row in delivered["recommendations"] if normalized(row["Area"]) not in ppt_norm]
    ppt_content_ok = 4 <= slide_count <= 8 and abs(aspect - 16 / 9) < 0.03 and not missing_papers and not missing_metrics and method_ok and not recommendation_areas_missing
    results.append(criterion("R009", ppt_content_ok, f"页数={slide_count}，比例={aspect:.3f}，缺论文信息={missing_papers}，缺指标={missing_metrics}，方法={method_ok}，缺建议领域={recommendation_areas_missing}"))

    empty_slides = [index for index, text in enumerate(slide_texts, start=1) if len(normalized(text)) < 40]
    ppt_native_ok = native_objects >= 25 and not oversized_pictures and not out_of_bounds and not empty_slides
    results.append(criterion("R010", ppt_native_ok, f"原生对象={native_objects}，图片={pictures}，整页大图={oversized_pictures}，越界={out_of_bounds}，空页={empty_slides}"))

    delivered_rows_canonical = [
        {key: (float(row[key]) if key in ("Citations", "Relevance_Score") else str(row[key]).strip()) for key in DATA_HEADERS}
        for row in delivered["data_analysis"]
    ]
    excel_rows_canonical = [
        {key: (float(row[key]) if key in ("Citations", "Relevance_Score") else str(row[key]).strip()) for key in DATA_HEADERS}
        for row in data_rows
    ] if data_rows else []
    cross_ok = (
        delivered_rows_canonical == excel_rows_canonical
        and json_metrics_ok
        and excel_metrics_ok
        and not missing_papers
        and not missing_metrics
        and recommendations_cross_ok
        and not recommendation_areas_missing
    )
    results.append(criterion("R011", cross_ok, f"JSON/Excel数据={delivered_rows_canonical == excel_rows_canonical}，JSON指标={json_metrics_ok}，Excel指标={excel_metrics_ok}，PPT论文缺失={missing_papers}，PPT指标缺失={missing_metrics}，建议一致={recommendations_cross_ok}"))

    xlsx_render_norm = normalized(rendered_texts.get(XLSX_NAME, ""))
    ppt_render_norm = normalized(rendered_texts.get(PPTX_NAME, ""))
    xlsx_render_ok = all(
        token in xlsx_render_norm
        for token in ("paper_id", "computer vision", "large language models", "total_papers", "priority", "action")
    )
    ppt_render_ok = rendered_pages.get(PPTX_NAME) == slide_count and all(token in ppt_render_norm for token in ("научные", "large language models", "avg_relevance", "gap"))
    render_ok = recalculated is not None and xlsx_render_ok and ppt_render_ok
    results.append(criterion("R012", render_ok, f"{render_evidence}，页数={rendered_pages}，Excel文本={xlsx_render_ok}，PPT文本={ppt_render_ok}"))

    formatting_ok = all(
        item.get("exists", True)
        and item.get("tables", 0) >= 1
        and item.get("freeze") == "A2"
        and item.get("header_bold")
        and item.get("header_fill")
        and item.get("wide_columns")
        for item in formatting.values()
    )
    results.append(criterion("R013", formatting_ok, json.dumps(formatting, ensure_ascii=False)))

    excel_text = "\n".join(str(cell) for row in data_rows + metric_rows + recommendation_rows for cell in row.values())
    unsupported = unsupported_gap_claims(json.dumps(delivered, ensure_ascii=False), excel_text, ppt_blob)
    source_error = not (json_rows_ok and json_metrics_ok and excel_rows_ok and excel_metrics_ok and cross_ok)
    script_error = not (canonical_ok and mutated_ok)
    broken_or_conflict = not (ppt_native_ok and render_ok and cross_ok and excel_structure_ok)
    results.append(criterion("R014", source_error, f"JSON行={json_rows_ok}，JSON指标={json_metrics_ok}，Excel行={excel_rows_ok}，Excel指标={excel_metrics_ok}，跨文件={cross_ok}"))
    results.append(criterion("R015", script_error, f"原始回放={canonical_ok}，变异回放={mutated_ok}"))
    results.append(criterion("R016", broken_or_conflict, f"PPT原生={ppt_native_ok}，渲染={render_ok}，跨文件={cross_ok}，Excel结构={excel_structure_ok}"))
    results.append(criterion("R017", bool(unsupported), f"无来源 gap/趋势声明={unsupported}"))
    return build_result(results, list(paths.values()))


def build_result(results: list[dict], artifact_paths: list[Path]) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付不存在、损坏或门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[identifier] for identifier in hurdle_failures))
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in artifact_paths if path.is_file()]
    if not digests:
        digests = [hashlib.sha256(b"").hexdigest()]
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": score,
                "evidence": item["evidence"],
                "method": item["method"],
                "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]},
            }
        )
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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
