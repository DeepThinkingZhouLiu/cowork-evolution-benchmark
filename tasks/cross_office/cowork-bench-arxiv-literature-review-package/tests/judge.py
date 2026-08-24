#!/usr/bin/env python3
"""对 Arxiv/API 静态输入、处理脚本、JSON、Excel 与 Word 做端到端确定性评分。"""

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
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


PROCESSOR_NAME = "arxiv_litreview_processor.py"
JSON_NAME = "arxiv_litreview_results.json"
REPORT_NAME = "Lit_Review_Report.xlsx"
DOCX_NAME = "Lit_Review_Analysis.docx"
TRACKER_NAME = "Arxiv_Litreview_Tracker.xlsx"
OUTPUT_NAMES = [PROCESSOR_NAME, JSON_NAME, REPORT_NAME, DOCX_NAME, TRACKER_NAME]
TASK_ID = "cowork-evo/cowork-bench-arxiv-literature-review-package"
RUBRIC_VERSION = "cowork-bench-arxiv-literature-review-package-v0.1"
JUDGE_VERSION = "cowork-bench-arxiv-literature-review-package-v0.1"
PASS_THRESHOLD = 0.8
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.35,
    "R003": 0.45,
    "R004": 0.4,
    "R005": 0.45,
    "R006": 0.5,
    "R007": 0.65,
    "R008": 0.55,
    "R009": 0.5,
    "R010": 0.65,
    "R011": 0.6,
}
DATA_HEADERS = [
    "Paper_ID",
    "Title",
    "Primary_Category",
    "Published",
    "Benchmark_Area",
    "Growth_Pct",
    "Match_Status",
]
EXCLUDED_HEADERS = ["Paper_ID", "Title", "Primary_Category", "Exclusion_Reason"]
METRIC_ORDER = [
    "Total_Papers",
    "Excluded_Noise_Papers",
    "Benchmark_Matched_Papers",
    "Avg_Matched_Growth_Pct",
    "Top_Growth_Area",
    "Benchmark_Coverage_Pct",
]
REC_HEADERS = ["Rank", "Priority", "Action", "Area", "Growth_Pct"]
TRACKER_HEADERS = ["Paper_ID", "Title", "Primary_Category", "Benchmark_Area", "Growth_Pct", "Review_Status"]
TOP_LEVEL_KEYS = ["sources", "source_summary", "data_analysis", "excluded_papers", "metrics", "recommendations", "limitations"]

CRITERION_SPECS = {
    "R001": ("5个指定交付物完整可读，Excel工作表正确", 6, "hurdle"),
    "R002": ("处理脚本在原始输入上按指定CLI回放成功并输出真值", 8, "hurdle"),
    "R003": ("处理脚本在Paper_ID和增长率变异输入上正确响应", 8, "hurdle"),
    "R004": ("交付JSON的来源摘要、主明细、排除项与指标精确正确", 8, "hurdle"),
    "R005": ("Report的Data_Analysis结构、排序和真值正确", 7, "hurdle"),
    "R006": ("Metrics的6项值正确且其中5项使用可审计公式", 7, "hurdle"),
    "R007": ("JSON和Report的3条建议排名、领域和增长率正确", 5, "hurdle"),
    "R008": ("离线Tracker的4条记录与JSON/Report一致", 6, "hurdle"),
    "R009": ("Word俄文报告结构、方法、核心事实和建议完整正确", 8, "hurdle"),
    "R010": ("交付物明确标注来源、匹配方法与数据边界", 5, "hurdle"),
    "R011": ("JSON、Report、Tracker和Word之间核心事实一致", 4, "hurdle"),
    "R012": ("Excel与Word保持原生可编辑结构和专业格式", 4, "positive"),
    "R013": ("两份Excel和Word可由LibreOffice重算、打开并渲染", 4, "positive"),
    "R014": ("任一来源派生核心事实与输入真值冲突", -18, "penalty"),
    "R015": ("处理脚本写死、复制现成结果或不随变异输入变化", -18, "penalty"),
    "R016": ("交付物缺失、损坏、截图化或跨文件核心事实冲突", -18, "penalty"),
    "R017": ("交付中存在无来源引用数、相关性、gap、时间序列或论文表现断言", -20, "penalty"),
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


def normalized(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def close(actual: Any, expected: Any, tolerance: float = 0.001) -> bool:
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return False


def round_half_up(value: float, digits: int = 2) -> float:
    quantum = Decimal("1").scaleb(-digits)
    return float(Decimal(str(value)).quantize(quantum, rounding=ROUND_HALF_UP))


def infer_area(record: dict[str, Any]) -> str | None:
    evidence = f"{record.get('title', '')} {record.get('summary', '')}".lower()
    for keyword, area in (
        ("language model", "LLMs"),
        ("vision transformer", "Computer Vision"),
        ("reinforcement learning", "RL"),
    ):
        if keyword in evidence:
            return area
    return None


def derive_expected(papers_path: Path, benchmark_path: Path) -> dict[str, Any]:
    papers = json.loads(papers_path.read_text(encoding="utf-8"))
    benchmark = json.loads(benchmark_path.read_text(encoding="utf-8"))
    if not isinstance(papers, list) or not papers:
        raise ValueError("论文JSON不是非空数组")
    topics = benchmark.get("trending_topics") if isinstance(benchmark, dict) else None
    if not isinstance(topics, list) or not topics:
        raise ValueError("benchmark缺少trending_topics")
    growth_by_area: dict[str, float] = {}
    for topic in topics:
        area = str(topic["area"])
        if area in growth_by_area:
            raise ValueError(f"重复benchmark area: {area}")
        growth_by_area[area] = float(topic["growth_pct"])

    included: list[dict[str, Any]] = []
    excluded: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    for paper in papers:
        paper_id = str(paper["id"])
        if paper_id in seen_ids:
            raise ValueError(f"重复Paper_ID: {paper_id}")
        seen_ids.add(paper_id)
        category = str(paper["primary_category"])
        if not category.startswith("cs."):
            excluded.append(
                {
                    "Paper_ID": paper_id,
                    "Title": str(paper["title"]),
                    "Primary_Category": category,
                    "Exclusion_Reason": "primary_category does not start with cs.",
                }
            )
            continue
        inferred = infer_area(paper)
        matched = inferred is not None and inferred in growth_by_area
        included.append(
            {
                "Paper_ID": paper_id,
                "Title": str(paper["title"]),
                "Primary_Category": category,
                "Published": str(paper["published"]),
                "Benchmark_Area": inferred if matched else None,
                "Growth_Pct": growth_by_area[inferred] if matched else None,
                "Match_Status": "Matched" if matched else "Unmatched",
            }
        )
    included.sort(key=lambda row: row["Paper_ID"])
    excluded.sort(key=lambda row: row["Paper_ID"])
    matched_rows = [row for row in included if row["Match_Status"] == "Matched"]
    ranked = sorted(matched_rows, key=lambda row: (-float(row["Growth_Pct"]), str(row["Benchmark_Area"])))
    growth = [float(row["Growth_Pct"]) for row in matched_rows]
    priorities = ["High", "Medium", "Monitor"]
    return {
        "source_summary": {
            "Input_Papers": len(papers),
            "Relevant_Papers": len(included),
            "Excluded_Papers": len(excluded),
            "Benchmark_Topics": len(topics),
        },
        "data_analysis": included,
        "excluded_papers": excluded,
        "metrics": {
            "Total_Papers": len(included),
            "Excluded_Noise_Papers": len(excluded),
            "Benchmark_Matched_Papers": len(matched_rows),
            "Avg_Matched_Growth_Pct": round_half_up(sum(growth) / len(growth), 2) if growth else None,
            "Top_Growth_Area": ranked[0]["Benchmark_Area"] if ranked else None,
            "Benchmark_Coverage_Pct": round_half_up(100 * len(matched_rows) / len(included), 2) if included else 0.0,
        },
        "recommendations": [
            {
                "Rank": index,
                "Priority": priorities[min(index - 1, len(priorities) - 1)],
                "Area": row["Benchmark_Area"],
                "Growth_Pct": float(row["Growth_Pct"]),
            }
            for index, row in enumerate(ranked, start=1)
        ],
    }


def load_result(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or list(data) != TOP_LEVEL_KEYS:
        raise ValueError(f"JSON顶层键错误: {list(data) if isinstance(data, dict) else type(data).__name__}")
    return data


def validate_rows(actual: Any, expected: list[dict[str, Any]], headers: list[str] = DATA_HEADERS) -> tuple[bool, list[Any]]:
    errors: list[Any] = []
    if not isinstance(actual, list) or len(actual) != len(expected):
        return False, [("row_count", len(actual) if isinstance(actual, list) else type(actual).__name__, len(expected))]
    for index, (actual_row, expected_row) in enumerate(zip(actual, expected)):
        if not isinstance(actual_row, dict) or list(actual_row) != headers:
            errors.append((index, "keys", list(actual_row) if isinstance(actual_row, dict) else type(actual_row).__name__, headers))
            continue
        for key in headers:
            actual_value = actual_row.get(key)
            expected_value = expected_row.get(key)
            if key == "Growth_Pct" and expected_value is not None:
                if not close(actual_value, expected_value):
                    errors.append((index, key, actual_value, expected_value))
            elif actual_value is None or expected_value is None:
                if actual_value not in (None, "") or expected_value not in (None, ""):
                    errors.append((index, key, actual_value, expected_value))
            elif str(actual_value).strip() != str(expected_value).strip():
                errors.append((index, key, actual_value, expected_value))
    return not errors, errors


def validate_metrics(actual: Any, expected: dict[str, Any]) -> tuple[bool, list[Any]]:
    errors: list[Any] = []
    if not isinstance(actual, dict) or list(actual) != METRIC_ORDER:
        return False, [("keys", list(actual) if isinstance(actual, dict) else type(actual).__name__, METRIC_ORDER)]
    for key in METRIC_ORDER:
        expected_value = expected[key]
        actual_value = actual.get(key)
        if isinstance(expected_value, str):
            if str(actual_value or "").strip() != expected_value:
                errors.append((key, actual_value, expected_value))
        elif expected_value is None:
            if actual_value is not None:
                errors.append((key, actual_value, expected_value))
        elif not close(actual_value, expected_value, 0.011):
            errors.append((key, actual_value, expected_value))
    return not errors, errors


def validate_recommendations(actual: Any, expected: list[dict[str, Any]]) -> tuple[bool, list[Any]]:
    errors: list[Any] = []
    if not isinstance(actual, list) or len(actual) != len(expected):
        return False, [("row_count", len(actual) if isinstance(actual, list) else type(actual).__name__, len(expected))]
    for index, (actual_row, expected_row) in enumerate(zip(actual, expected)):
        if not isinstance(actual_row, dict) or list(actual_row) != REC_HEADERS:
            errors.append((index, "keys", list(actual_row) if isinstance(actual_row, dict) else type(actual_row).__name__, REC_HEADERS))
            continue
        for key in ("Rank", "Priority", "Area"):
            if str(actual_row.get(key) or "").strip() != str(expected_row[key]):
                errors.append((index, key, actual_row.get(key), expected_row[key]))
        if not str(actual_row.get("Action") or "").strip():
            errors.append((index, "Action", "blank"))
        if not close(actual_row.get("Growth_Pct"), expected_row["Growth_Pct"]):
            errors.append((index, "Growth_Pct", actual_row.get("Growth_Pct"), expected_row["Growth_Pct"]))
    return not errors, errors


def validate_result(actual: dict[str, Any] | None, expected: dict[str, Any]) -> tuple[bool, dict[str, Any]]:
    if actual is None:
        return False, {"result": "missing"}
    source_summary_ok = actual.get("source_summary") == expected["source_summary"]
    rows_ok, row_errors = validate_rows(actual.get("data_analysis"), expected["data_analysis"])
    excluded_ok, excluded_errors = validate_rows(actual.get("excluded_papers"), expected["excluded_papers"], EXCLUDED_HEADERS)
    metrics_ok, metric_errors = validate_metrics(actual.get("metrics"), expected["metrics"])
    recs_ok, rec_errors = validate_recommendations(actual.get("recommendations"), expected["recommendations"])
    sources = actual.get("sources")
    sources_ok = isinstance(sources, list) and {row.get("file") for row in sources if isinstance(row, dict)} == {"arxiv_papers.json", "benchmark_data.json"}
    limitations = actual.get("limitations")
    limitations_ok = isinstance(limitations, list) and len(limitations) >= 4 and all(str(item).strip() for item in limitations)
    evidence = {
        "source_summary": source_summary_ok,
        "sources": sources_ok,
        "rows": row_errors,
        "excluded": excluded_errors,
        "metrics": metric_errors,
        "recommendations": rec_errors,
        "limitations": limitations_ok,
    }
    return all((source_summary_ok, sources_ok, rows_ok, excluded_ok, metrics_ok, recs_ok, limitations_ok)), evidence


def run_processor(script: Path, papers: Path, benchmark: Path, output: Path) -> tuple[dict[str, Any] | None, str]:
    try:
        completed = subprocess.run(
            [
                sys.executable,
                str(script.resolve()),
                "--papers",
                str(papers.resolve()),
                "--benchmark",
                str(benchmark.resolve()),
                "--output",
                str(output.resolve()),
            ],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
            cwd=output.parent,
        )
        if completed.returncode != 0 or not output.is_file():
            return None, f"exit={completed.returncode}, stdout={completed.stdout[-500:]}, stderr={completed.stderr[-500:]}"
        return load_result(output), "exit=0, JSON可读"
    except Exception as error:
        return None, str(error)


def read_sheet_rows(sheet, headers: list[str]) -> tuple[list[dict[str, Any]], list[str]]:
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


def formula_norm(value: Any) -> str:
    return re.sub(r"[\s$']", "", str(value or "")).upper()


def metric_formulas_ok(workbook) -> tuple[bool, dict[str, str]]:
    sheet = workbook["Metrics"]
    formulas = {f"B{row}": str(sheet[f"B{row}"].value or "") for row in range(2, 8)}
    values = {cell: formula_norm(value) for cell, value in formulas.items()}
    checks = [
        values["B2"].startswith("=") and "COUNTA" in values["B2"] and "DATA_ANALYSIS!A2:A5" in values["B2"],
        not values["B3"].startswith("=") and close(formulas["B3"], 1),
        values["B4"].startswith("=") and "COUNT" in values["B4"] and "DATA_ANALYSIS!F2:F5" in values["B4"],
        values["B5"].startswith("=") and "AVERAGE" in values["B5"] and "DATA_ANALYSIS!F2:F5" in values["B5"],
        values["B6"].startswith("=") and "INDEX" in values["B6"] and "MAX" in values["B6"] and "DATA_ANALYSIS!E2:E5" in values["B6"] and "DATA_ANALYSIS!F2:F5" in values["B6"],
        values["B7"].startswith("=") and "COUNT" in values["B7"] and "COUNTA" in values["B7"] and "DATA_ANALYSIS!F2:F5" in values["B7"] and "DATA_ANALYSIS!A2:A5" in values["B7"],
    ]
    return all(checks), formulas


def docx_text_and_structure(path: Path) -> tuple[str, list[str], int, int]:
    document = Document(path)
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    headings = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.style.name.startswith("Heading")]
    inline_shapes = len(document.inline_shapes)
    return "\n".join(parts), headings, len(document.tables), inline_shapes


def office_formatting(report_book, tracker_book, docx_path: Path) -> tuple[bool, dict[str, Any]]:
    details: dict[str, Any] = {}
    for name, book, expected_sheets in (
        ("report", report_book, ["Data_Analysis", "Metrics", "Recommendations"]),
        ("tracker", tracker_book, ["Arxiv Litreview Tracker"]),
    ):
        sheet_details = {}
        for sheet_name in expected_sheets:
            if sheet_name not in book.sheetnames:
                sheet_details[sheet_name] = {"exists": False}
                continue
            sheet = book[sheet_name]
            header_cells = list(sheet[1])
            widths = [sheet.column_dimensions[cell.column_letter].width or 0 for cell in header_cells]
            sheet_details[sheet_name] = {
                "exists": True,
                "tables": len(sheet.tables),
                "freeze": str(sheet.freeze_panes or ""),
                "header_bold": all(bool(cell.font.bold) for cell in header_cells),
                "header_fill": all(cell.fill.fill_type == "solid" for cell in header_cells),
                "widths_ok": all(width >= 9 for width in widths),
                "nonempty_cells": sum(1 for row in sheet.iter_rows() for cell in row if cell.value is not None),
            }
        details[name] = sheet_details
    try:
        _, headings, table_count, inline_shapes = docx_text_and_structure(docx_path)
        details["word"] = {"headings": headings, "tables": table_count, "inline_shapes": inline_shapes}
    except Exception as error:
        details["word"] = {"error": str(error)}
    excel_ok = all(
        item.get("exists")
        and item.get("tables", 0) >= 1
        and item.get("freeze") == "A2"
        and item.get("header_bold")
        and item.get("header_fill")
        and item.get("widths_ok")
        and item.get("nonempty_cells", 0) >= 8
        for group in (details["report"], details["tracker"])
        for item in group.values()
    )
    word = details.get("word", {})
    word_ok = len(word.get("headings", [])) >= 5 and word.get("tables", 0) >= 2 and word.get("inline_shapes", 0) == 0
    return excel_ok and word_ok, details


def libreoffice_prepare(report: Path, tracker: Path, docx: Path, root: Path) -> tuple[Path | None, dict[str, str], dict[str, int], str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    pdfinfo = shutil.which("pdfinfo")
    if libreoffice is None or pdftotext is None or pdfinfo is None:
        return None, {}, {}, "缺少libreoffice/pdftotext/pdfinfo"
    source_dir = root / "source"
    recalc_dir = root / "recalc"
    pdf_dir = root / "pdf"
    source_dir.mkdir()
    recalc_dir.mkdir()
    pdf_dir.mkdir()
    copies = {}
    for source in (report, tracker, docx):
        target = source_dir / source.name
        shutil.copy2(source, target)
        copies[source.name] = target
    profile_uri = (root / "lo-profile").resolve().as_uri()
    recalc = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(copies[REPORT_NAME])],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    recalculated = recalc_dir / REPORT_NAME
    if recalc.returncode != 0 or not recalculated.is_file():
        return None, {}, {}, (recalc.stderr or recalc.stdout or "Excel重算失败")[-1000:]
    texts: dict[str, str] = {}
    pages: dict[str, int] = {}
    render_sources = [recalculated, copies[TRACKER_NAME], copies[DOCX_NAME]]
    for source in render_sources:
        converted = subprocess.run(
            [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(source)],
            capture_output=True,
            text=True,
            timeout=300,
            check=False,
        )
        pdf_path = pdf_dir / f"{source.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return recalculated, texts, pages, (converted.stderr or converted.stdout or f"{source.name}渲染失败")[-1000:]
        extracted = subprocess.run([pdftotext, "-raw", str(pdf_path), "-"], capture_output=True, text=True, timeout=120, check=False)
        info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=120, check=False)
        if extracted.returncode != 0 or info.returncode != 0:
            return recalculated, texts, pages, f"{source.name} PDF检查失败"
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        texts[source.name] = extracted.stdout
        pages[source.name] = int(match.group(1)) if match else 0
    return recalculated, texts, pages, "Report重算且Report/Tracker/Word渲染成功"


def unsupported_assertions(delivered: dict[str, Any], report_headers: list[str], tracker_headers: list[str], office_text: str) -> list[str]:
    findings: list[str] = []
    serialized = json.dumps(delivered, ensure_ascii=False)

    def collect_keys(value: Any) -> set[str]:
        keys: set[str] = set()
        if isinstance(value, dict):
            keys.update(str(key).lower() for key in value)
            for child in value.values():
                keys.update(collect_keys(child))
        elif isinstance(value, list):
            for child in value:
                keys.update(collect_keys(child))
        return keys

    keys = collect_keys(delivered)
    all_headers = {normalized(value).replace(" ", "_") for value in report_headers + tracker_headers}
    if "citations" in keys or "citations" in all_headers:
        findings.append("unsupported_citations_field")
    if "relevance_score" in keys or "relevance_score" in all_headers:
        findings.append("unsupported_relevance_field")

    combined = f"{serialized}\n{office_text}"
    patterns = {
        "upstream_invented_citations": r"(?:citation|citations|цитировани\w*|цитат\w*)\D{0,30}(?:450|320|180|250|300)\b",
        "upstream_invented_relevance": r"(?:relevance[_ ]?score|оценк\w* релевантност\w*)\D{0,30}(?:95|82|71|68)\b",
        "numeric_gap": r"(?:gap|gap-анализ|разрыв|отклонени\w*)\s*(?:=|:|составляет)?\s*-?\d+(?:[.,]\d+)?\s*%",
        "unsupported_period": r"(?:Q1\s*2026|1[- ]?й квартал 2026|I квартал 2026).{0,50}(?:данные|анализ|рост)",
        "unsupported_correlation": r"(?:r\s*=\s*0[.,]\d+|correlat\w*\s*(?:=|of)?\s*0[.,]\d+|корреляц\w*\s*(?:=|составляет)?\s*0[.,]\d+)",
    }
    for label, pattern in patterns.items():
        if re.search(pattern, combined, re.IGNORECASE | re.DOTALL):
            findings.append(label)
    return findings


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results: list[dict[str, Any]] = []
    missing = [name for name, path in paths.items() if not path.is_file()]
    if missing:
        results.append(criterion("R001", False, f"缺失={missing}"))
        results.append(criterion("R016", True, f"缺失={missing}"))
        return build_result(results, list(paths.values()))

    input_dir = Path(__file__).resolve().parents[1] / "data" / "input_files"
    papers_path = input_dir / "arxiv_papers.json"
    benchmark_path = input_dir / "benchmark_data.json"
    expected = derive_expected(papers_path, benchmark_path)
    try:
        delivered = load_result(paths[JSON_NAME])
        if not zipfile.is_zipfile(paths[REPORT_NAME]) or not zipfile.is_zipfile(paths[TRACKER_NAME]) or not zipfile.is_zipfile(paths[DOCX_NAME]):
            raise ValueError("Office文件不是有效OOXML ZIP")
        report_formula_book = load_workbook(paths[REPORT_NAME], data_only=False)
        tracker_book = load_workbook(paths[TRACKER_NAME], data_only=False)
        if report_formula_book.sheetnames != ["Data_Analysis", "Metrics", "Recommendations"]:
            raise ValueError(f"Report工作表错误: {report_formula_book.sheetnames}")
        if tracker_book.sheetnames != ["Arxiv Litreview Tracker"]:
            raise ValueError(f"Tracker工作表错误: {tracker_book.sheetnames}")
        script_text = paths[PROCESSOR_NAME].read_text(encoding="utf-8")
        if len(script_text.strip()) < 500:
            raise ValueError("处理脚本过短")
        word_text, word_headings, word_tables, word_inline_shapes = docx_text_and_structure(paths[DOCX_NAME])
    except Exception as error:
        results.append(criterion("R001", False, f"交付读取失败: {error}"))
        results.append(criterion("R016", True, f"交付读取失败: {error}"))
        return build_result(results, list(paths.values()))
    results.append(criterion("R001", True, f"5个文件可读；Report={report_formula_book.sheetnames}；Tracker={tracker_book.sheetnames}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-arxiv-script-") as temp_dir:
        temp = Path(temp_dir)
        canonical, canonical_run = run_processor(paths[PROCESSOR_NAME], papers_path, benchmark_path, temp / "canonical.json")
        canonical_ok, canonical_details = validate_result(canonical, expected)

        mutated_papers = json.loads(papers_path.read_text(encoding="utf-8"))
        mutated_benchmark = json.loads(benchmark_path.read_text(encoding="utf-8"))
        mutated_papers[1]["id"] = "2402.05678"
        for topic in mutated_benchmark["trending_topics"]:
            if topic["area"] == "LLMs":
                topic["growth_pct"] = 55
            elif topic["area"] == "Computer Vision":
                topic["growth_pct"] = 22
        mutated_papers_path = temp / "arxiv_papers_mutated.json"
        mutated_benchmark_path = temp / "benchmark_data_mutated.json"
        mutated_papers_path.write_text(json.dumps(mutated_papers, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        mutated_benchmark_path.write_text(json.dumps(mutated_benchmark, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        mutated_expected = derive_expected(mutated_papers_path, mutated_benchmark_path)
        mutated, mutated_run = run_processor(
            paths[PROCESSOR_NAME], mutated_papers_path, mutated_benchmark_path, temp / "mutated.json"
        )
        mutated_valid, mutated_details = validate_result(mutated, mutated_expected)
        mutation_changed = (
            canonical is not None
            and mutated is not None
            and canonical.get("data_analysis") != mutated.get("data_analysis")
            and canonical.get("metrics") != mutated.get("metrics")
            and any(row.get("Paper_ID") == "2402.05678" for row in mutated.get("data_analysis", []))
            and close(mutated.get("metrics", {}).get("Avg_Matched_Growth_Pct"), 28.33, 0.011)
        )
    results.append(criterion("R002", canonical_ok, f"{canonical_run}；详情={canonical_details}"))
    results.append(criterion("R003", mutated_valid and mutation_changed, f"{mutated_run}；详情={mutated_details}；已变化={mutation_changed}；变异指标={mutated_expected['metrics']}"))

    delivered_ok, delivered_details = validate_result(delivered, expected)
    results.append(criterion("R004", delivered_ok, json.dumps(delivered_details, ensure_ascii=False)))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-arxiv-office-") as temp_dir:
        recalculated, rendered_texts, rendered_pages, render_evidence = libreoffice_prepare(
            paths[REPORT_NAME], paths[TRACKER_NAME], paths[DOCX_NAME], Path(temp_dir)
        )
        value_report_path = recalculated or paths[REPORT_NAME]
        try:
            report_value_book = load_workbook(value_report_path, data_only=True)
            data_rows, data_missing = read_sheet_rows(report_value_book["Data_Analysis"], DATA_HEADERS)
            metric_rows, metric_missing = read_sheet_rows(report_value_book["Metrics"], ["Metric", "Value"])
            rec_rows, rec_missing = read_sheet_rows(report_value_book["Recommendations"], REC_HEADERS)
            tracker_rows, tracker_missing = read_sheet_rows(tracker_book["Arxiv Litreview Tracker"], TRACKER_HEADERS)
            formula_ok, formula_details = metric_formulas_ok(report_formula_book)
            formatting_ok, formatting_details = office_formatting(report_formula_book, tracker_book, paths[DOCX_NAME])
        except Exception as error:
            results.append(criterion("R005", False, f"Excel读取失败: {error}"))
            results.append(criterion("R016", True, f"Excel读取失败: {error}"))
            return build_result(results, list(paths.values()))

    excel_rows_ok, excel_row_errors = validate_rows(data_rows, expected["data_analysis"])
    data_sorted = [str(row.get("Paper_ID") or "") for row in data_rows] == sorted(str(row.get("Paper_ID") or "") for row in data_rows)
    report_structure_ok = not data_missing and not metric_missing and not rec_missing
    results.append(criterion("R005", report_structure_ok and excel_rows_ok and data_sorted, f"缺表头={data_missing + metric_missing + rec_missing}；数据错={excel_row_errors}；已排序={data_sorted}"))

    actual_metrics = {str(row["Metric"] or "").strip(): row["Value"] for row in metric_rows}
    excel_metric_errors = []
    if list(actual_metrics) != METRIC_ORDER:
        excel_metric_errors.append(("order", list(actual_metrics), METRIC_ORDER))
    for key in METRIC_ORDER:
        if key not in actual_metrics:
            excel_metric_errors.append((key, "missing", expected["metrics"][key]))
        elif isinstance(expected["metrics"][key], str):
            if str(actual_metrics[key] or "").strip() != expected["metrics"][key]:
                excel_metric_errors.append((key, actual_metrics[key], expected["metrics"][key]))
        else:
            comparison_value = actual_metrics[key]
            # Excel 百分比单元格以 0.75 存储并显示为 75.00%；JSON 用 75.0 表示百分点。
            if key == "Benchmark_Coverage_Pct" and close(expected["metrics"][key], 75.0, 0.011):
                try:
                    if abs(float(comparison_value)) <= 1:
                        comparison_value = float(comparison_value) * 100
                except (TypeError, ValueError):
                    pass
            if not close(comparison_value, expected["metrics"][key], 0.011):
                excel_metric_errors.append((key, actual_metrics[key], expected["metrics"][key]))
    excel_metrics_ok = not excel_metric_errors
    results.append(criterion("R006", formula_ok and excel_metrics_ok and recalculated is not None, f"公式={formula_details}；值错={excel_metric_errors}；重算={recalculated is not None}"))

    excel_recs = []
    for row in rec_rows:
        excel_recs.append({key: row.get(key) for key in REC_HEADERS})
    excel_recs_ok, excel_rec_errors = validate_recommendations(excel_recs, expected["recommendations"])
    json_recs_ok, json_rec_errors = validate_recommendations(delivered.get("recommendations"), expected["recommendations"])
    recs_cross = [
        (row.get("Rank"), row.get("Priority"), row.get("Area"), float(row.get("Growth_Pct")))
        for row in excel_recs
    ] == [
        (row.get("Rank"), row.get("Priority"), row.get("Area"), float(row.get("Growth_Pct")))
        for row in delivered.get("recommendations", [])
    ] if excel_recs and delivered.get("recommendations") else False
    results.append(criterion("R007", excel_recs_ok and json_recs_ok and recs_cross, f"Excel错={excel_rec_errors}；JSON错={json_rec_errors}；跨文件={recs_cross}"))

    expected_tracker = []
    for row in expected["data_analysis"]:
        expected_tracker.append(
            {
                "Paper_ID": row["Paper_ID"],
                "Title": row["Title"],
                "Primary_Category": row["Primary_Category"],
                "Benchmark_Area": row["Benchmark_Area"],
                "Growth_Pct": row["Growth_Pct"],
                "Review_Status": "Ready for review" if row["Match_Status"] == "Matched" else "Needs separate benchmark",
            }
        )
    tracker_ok, tracker_errors = validate_rows(tracker_rows, expected_tracker, TRACKER_HEADERS)
    results.append(criterion("R008", not tracker_missing and tracker_ok, f"缺表头={tracker_missing}；错误={tracker_errors}"))

    word_norm = normalized(word_text)
    required_headings = ["краткое резюме", "источники и метод", "ключевые выводы", "рекомендации", "ограничения"]
    heading_norm = [normalized(value) for value in word_headings]
    missing_headings = [value for value in required_headings if value not in heading_norm]
    required_word_tokens = [
        "arxiv_papers.json",
        "benchmark_data.json",
        "5",
        "4",
        "3",
        "21,67%",
        "75,00%",
        "llms",
        "45%",
        "computer vision",
        "12%",
        "rl",
        "8%",
        "quant-ph",
        "federated learning",
    ] + [row["Paper_ID"] for row in expected["data_analysis"]]
    missing_word_tokens = [token for token in required_word_tokens if normalized(token) not in word_norm]
    cyrillic_count = len(re.findall(r"[А-Яа-яЁё]", word_text))
    word_ok = len(word_text) >= 1800 and cyrillic_count >= 1000 and not missing_headings and not missing_word_tokens and word_tables >= 2
    results.append(criterion("R009", word_ok, f"字符={len(word_text)}；西里尔={cyrillic_count}；缺标题={missing_headings}；缺事实={missing_word_tokens}；表格={word_tables}"))

    json_blob = normalized(json.dumps(delivered, ensure_ascii=False))
    source_ok = all(token in json_blob and token in word_norm for token in ("arxiv_papers.json", "benchmark_data.json"))
    method_ok = all(token in word_norm for token in ("primary_category", "language model", "vision transformer", "reinforcement learning"))
    limitation_ok = (
        isinstance(delivered.get("limitations"), list)
        and len(delivered["limitations"]) >= 4
        and all(token in word_norm for token in ("нет счётчиков цитирования", "gap-анализ", "нет временного ряда", "не измерение качества"))
    )
    results.append(criterion("R010", source_ok and method_ok and limitation_ok, f"来源={source_ok}；方法={method_ok}；限制={limitation_ok}"))

    json_rows_ok, _ = validate_rows(delivered.get("data_analysis"), expected["data_analysis"])
    json_metrics_ok, _ = validate_metrics(delivered.get("metrics"), expected["metrics"])
    word_cross_ok = not missing_word_tokens
    cross_ok = json_rows_ok and json_metrics_ok and excel_rows_ok and excel_metrics_ok and tracker_ok and recs_cross and word_cross_ok
    results.append(criterion("R011", cross_ok, f"JSON行={json_rows_ok}；JSON指标={json_metrics_ok}；Report行={excel_rows_ok}；Report指标={excel_metrics_ok}；Tracker={tracker_ok}；建议={recs_cross}；Word={word_cross_ok}"))

    results.append(criterion("R012", formatting_ok, json.dumps(formatting_details, ensure_ascii=False)))

    report_render = normalized(rendered_texts.get(REPORT_NAME, ""))
    tracker_render = normalized(rendered_texts.get(TRACKER_NAME, ""))
    word_render = normalized(rendered_texts.get(DOCX_NAME, ""))
    render_ok = (
        recalculated is not None
        and rendered_pages.get(REPORT_NAME, 0) >= 3
        and rendered_pages.get(TRACKER_NAME, 0) >= 1
        and rendered_pages.get(DOCX_NAME, 0) >= 2
        and all(token in report_render for token in ("paper_id", "2301.01234", "total_papers", "rank", "priority", "action"))
        and all(token in tracker_render for token in ("paper_id", "2304.03456", "review_status"))
        and all(token in word_render for token in ("краткое резюме", "llms", "45%", "ограничения"))
    )
    results.append(criterion("R013", render_ok, f"{render_evidence}；页数={rendered_pages}；Report文本={len(report_render)}；Tracker文本={len(tracker_render)}；Word文本={len(word_render)}"))

    report_headers = [str(cell.value or "") for sheet in report_formula_book.worksheets for cell in sheet[1]]
    tracker_headers = [str(cell.value or "") for cell in tracker_book["Arxiv Litreview Tracker"][1]]
    excel_text = "\n".join(str(value) for row in data_rows + metric_rows + rec_rows + tracker_rows for value in row.values())
    unsupported = unsupported_assertions(delivered, report_headers, tracker_headers, f"{excel_text}\n{word_text}")
    source_error = not (delivered_ok and excel_rows_ok and excel_metrics_ok and tracker_ok)
    script_error = not (canonical_ok and mutated_valid and mutation_changed)
    broken_or_conflict = not (formatting_ok and render_ok and cross_ok)
    results.append(criterion("R014", source_error, f"JSON={delivered_ok}；Report行={excel_rows_ok}；Report指标={excel_metrics_ok}；Tracker={tracker_ok}"))
    results.append(criterion("R015", script_error, f"原始回放={canonical_ok}；变异回放={mutated_valid}；已响应={mutation_changed}"))
    results.append(criterion("R016", broken_or_conflict, f"格式={formatting_ok}；渲染={render_ok}；跨文件={cross_ok}"))
    results.append(criterion("R017", bool(unsupported), f"无来源断言={unsupported}"))

    report_value_book.close()
    report_formula_book.close()
    tracker_book.close()
    return build_result(results, list(paths.values()))


def build_result(results: list[dict[str, Any]], artifact_paths: list[Path]) -> dict[str, Any]:
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
                "raw": {
                    "description": item["description"],
                    "weight": item["weight"],
                    "type": item["type"],
                    "passed": item["passed"],
                },
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
