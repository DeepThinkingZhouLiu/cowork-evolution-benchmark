#!/usr/bin/env python3
"""Arxiv 文献综述包的 v2 连续评分入口；v1 严格口径保存在 judge_v1.py。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Iterable

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-arxiv-literature-review-package"
RUBRIC_VERSION = "cowork-bench-arxiv-literature-review-package-v0.2"
JUDGE_VERSION = "cowork-bench-arxiv-literature-review-criterion-partial-v2.0"
PROCESSOR = "arxiv_litreview_processor.py"
RESULT_JSON = "arxiv_litreview_results.json"
REPORT = "Lit_Review_Report.xlsx"
WORD = "Lit_Review_Analysis.docx"
TRACKER = "Arxiv_Litreview_Tracker.xlsx"
REQUIRED = (PROCESSOR, RESULT_JSON, REPORT, WORD, TRACKER)
POSITIVE_WEIGHTS = {
    "R001": 6,
    "R002": 8,
    "R003": 8,
    "R004": 8,
    "R005": 7,
    "R006": 7,
    "R007": 5,
    "R008": 6,
    "R009": 8,
    "R010": 5,
    "R011": 4,
    "R012": 4,
    "R013": 4,
}
PENALTY_WEIGHTS = {"R014": -18, "R015": -18, "R016": -18, "R017": -20}


def _load_v1():
    path = Path(__file__).with_name("judge_v1.py")
    if str(path.parent) not in sys.path:
        sys.path.insert(0, str(path.parent))
    spec = importlib.util.spec_from_file_location("arxiv_litreview_judge_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入 v1 Judge")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


V1 = _load_v1()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _criterion(identifier: str, score: float, evidence: str, earned: float = 0, total: float = 1) -> dict[str, Any]:
    return {
        "criterion_id": identifier,
        "score": round(max(0.0, min(1.0, score)), 6),
        "evidence": evidence,
        "method": "deterministic",
        "raw": {"earned_units": round(earned, 6), "total_units": round(total, 6)},
    }


def _ratio(findings: Iterable[tuple[bool, str]]) -> tuple[float, list[str], int, int]:
    materialized = list(findings)
    earned = sum(bool(ok) for ok, _ in materialized)
    return earned / max(1, len(materialized)), [label for ok, label in materialized if not ok], earned, len(materialized)


def _same(actual: Any, expected: Any) -> bool:
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return V1.close(actual, expected, 0.011)
    if expected is None:
        return actual in (None, "")
    return str(actual).strip() == str(expected).strip()


def _mapping_findings(actual: Any, expected: dict[str, Any], label: str) -> list[tuple[bool, str]]:
    source = actual if isinstance(actual, dict) else {}
    findings = [(isinstance(actual, dict), f"{label} 应为对象")]
    for key, value in expected.items():
        findings.append((_same(source.get(key), value), f"{label}.{key} 错误"))
    return findings


def _rows_findings(actual: Any, expected: list[dict[str, Any]], headers: list[str], label: str) -> list[tuple[bool, str]]:
    rows = actual if isinstance(actual, list) else []
    findings = [(isinstance(actual, list), f"{label} 应为数组"), (len(rows) == len(expected), f"{label} 行数错误")]
    for row_index, wanted in enumerate(expected):
        seen = rows[row_index] if row_index < len(rows) and isinstance(rows[row_index], dict) else {}
        findings.append((list(seen) == headers, f"{label} 第 {row_index + 1} 行字段顺序错误"))
        for key in headers:
            findings.append((_same(seen.get(key), wanted.get(key)), f"{label} 第 {row_index + 1} 行 {key} 错误"))
    return findings


def _recommendation_findings(actual: Any, expected: list[dict[str, Any]], label: str) -> list[tuple[bool, str]]:
    rows = actual if isinstance(actual, list) else []
    findings = [(isinstance(actual, list), f"{label} 应为数组"), (len(rows) == 3, f"{label} 应有 3 行")]
    for index, wanted in enumerate(expected):
        seen = rows[index] if index < len(rows) and isinstance(rows[index], dict) else {}
        for key in ("Rank", "Priority", "Area", "Growth_Pct"):
            findings.append((_same(seen.get(key), wanted.get(key)), f"{label} 第 {index + 1} 行 {key} 错误"))
        findings.append((bool(str(seen.get("Action") or "").strip()), f"{label} 第 {index + 1} 行 Action 为空"))
    return findings


def _task_inputs() -> tuple[Path, Path]:
    task_root = Path(__file__).resolve().parents[1]
    input_dir = Path("/app/input_files") if Path("/app/input_files").is_dir() else task_root / "data" / "input_files"
    return input_dir / "arxiv_papers.json", input_dir / "benchmark_data.json"


def _core(artifact: Path) -> tuple[bool, list[str]]:
    if not artifact.is_dir():
        return False, ["输出目录不存在"]
    errors: list[str] = []
    for name in REQUIRED:
        path = artifact / name
        if not path.is_file() or path.stat().st_size == 0:
            errors.append(f"核心交付缺失或为空：{name}")
            continue
        try:
            if name == PROCESSOR:
                text = path.read_text(encoding="utf-8")
                compile(text, str(path), "exec")
                if len(text.strip()) < 100:
                    raise ValueError("处理脚本过短")
            elif name == RESULT_JSON:
                if not isinstance(json.loads(path.read_text(encoding="utf-8")), dict):
                    raise ValueError("JSON 顶层不是对象")
            elif name == REPORT:
                book = load_workbook(path, read_only=True, data_only=False)
                if book.sheetnames != ["Data_Analysis", "Metrics", "Recommendations"]:
                    raise ValueError(f"工作表错误：{book.sheetnames}")
                book.close()
            elif name == TRACKER:
                book = load_workbook(path, read_only=True, data_only=False)
                if book.sheetnames != ["Arxiv Litreview Tracker"]:
                    raise ValueError(f"工作表错误：{book.sheetnames}")
                book.close()
            elif name == WORD:
                document = Document(path)
                if not document.paragraphs:
                    raise ValueError("Word 没有段落")
        except Exception as exc:
            errors.append(f"核心交付无法解析：{name}: {type(exc).__name__}: {exc}")
    return not errors, errors


def _fatal(errors: list[str], hashes: list[str]) -> dict[str, Any]:
    identifiers = [*POSITIVE_WEIGHTS, *PENALTY_WEIGHTS]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": hashes or ["0" * 64],
        "criterion_results": [_criterion(key, 0, "；".join(errors), 0, 0) for key in identifiers],
        "reward": 0.0,
        "passed": False,
        "hurdle_failures": ["R001"],
        "penalties": [],
    }


def _result_findings(actual: dict[str, Any], expected: dict[str, Any]) -> list[tuple[bool, str]]:
    findings: list[tuple[bool, str]] = [(list(actual) == V1.TOP_LEVEL_KEYS, "JSON 顶层键或顺序错误")]
    sources = actual.get("sources")
    source_files = {item.get("file") for item in sources if isinstance(item, dict)} if isinstance(sources, list) else set()
    findings.extend([(isinstance(sources, list), "sources 应为数组"), (source_files == {"arxiv_papers.json", "benchmark_data.json"}, "sources 文件集合错误")])
    findings.extend(_mapping_findings(actual.get("source_summary"), expected["source_summary"], "source_summary"))
    findings.extend(_rows_findings(actual.get("data_analysis"), expected["data_analysis"], V1.DATA_HEADERS, "data_analysis"))
    findings.extend(_rows_findings(actual.get("excluded_papers"), expected["excluded_papers"], V1.EXCLUDED_HEADERS, "excluded_papers"))
    findings.extend(_mapping_findings(actual.get("metrics"), expected["metrics"], "metrics"))
    findings.extend(_recommendation_findings(actual.get("recommendations"), expected["recommendations"], "recommendations"))
    limitations = actual.get("limitations")
    findings.extend([(isinstance(limitations, list), "limitations 应为数组"), (isinstance(limitations, list) and len(limitations) >= 4, "limitations 少于 4 条"), (isinstance(limitations, list) and all(str(item).strip() for item in limitations), "limitations 含空项")])
    return findings


def _processor_checks(artifact: Path, mutated: bool) -> tuple[float, list[str], int, int, dict[str, Any] | None]:
    papers_path, benchmark_path = _task_inputs()
    with tempfile.TemporaryDirectory(prefix="cowork-evo-arxiv-v2-") as raw:
        temp = Path(raw)
        if mutated:
            papers = json.loads(papers_path.read_text(encoding="utf-8"))
            benchmark = json.loads(benchmark_path.read_text(encoding="utf-8"))
            papers[1]["id"] = "2402.05678"
            for topic in benchmark["trending_topics"]:
                if topic["area"] == "LLMs":
                    topic["growth_pct"] = 55
                elif topic["area"] == "Computer Vision":
                    topic["growth_pct"] = 22
            papers_path = temp / "arxiv_papers_mutated.json"
            benchmark_path = temp / "benchmark_data_mutated.json"
            papers_path.write_text(json.dumps(papers, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            benchmark_path.write_text(json.dumps(benchmark, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        expected = V1.derive_expected(papers_path, benchmark_path)
        actual, run_evidence = V1.run_processor(artifact / PROCESSOR, papers_path, benchmark_path, temp / "result.json")
        findings = [(actual is not None, f"处理脚本运行失败：{run_evidence}")]
        if actual is not None:
            findings.extend(_result_findings(actual, expected))
            if mutated:
                rows = actual.get("data_analysis") if isinstance(actual, dict) else []
                metrics = actual.get("metrics") if isinstance(actual, dict) else {}
                findings.extend([
                    (any(isinstance(row, dict) and row.get("Paper_ID") == "2402.05678" for row in rows or []), "变异 Paper_ID 未传播"),
                    (V1.close((metrics or {}).get("Avg_Matched_Growth_Pct"), 28.33, 0.011), "变异平均增长率不是 28.33"),
                    ((metrics or {}).get("Top_Growth_Area") == "LLMs", "变异最高增长领域错误"),
                ])
        score, failures, earned, total = _ratio(findings)
        return score, failures, earned, total, actual


def _report_rows(artifact: Path, expected: dict[str, Any]) -> tuple[float, list[str], int, int, dict[str, Any]]:
    book = load_workbook(artifact / REPORT, data_only=True)
    rows, missing = V1.read_sheet_rows(book["Data_Analysis"], V1.DATA_HEADERS)
    book.close()
    findings = [(not missing, f"Data_Analysis 缺表头：{missing}")]
    findings.extend(_rows_findings(rows, expected["data_analysis"], V1.DATA_HEADERS, "Report.Data_Analysis"))
    findings.append(([str(row.get("Paper_ID") or "") for row in rows] == sorted(str(row.get("Paper_ID") or "") for row in rows), "Data_Analysis 未按 Paper_ID 排序"))
    score, failures, earned, total = _ratio(findings)
    return score, failures, earned, total, {str(row.get("Paper_ID")): row for row in rows}


def _metrics(artifact: Path, expected: dict[str, Any]) -> tuple[float, list[str], int, int, dict[str, Any]]:
    formula_book = load_workbook(artifact / REPORT, data_only=False)
    formula_sheet = formula_book["Metrics"]
    formula_values = {f"B{row}": formula_sheet[f"B{row}"].value for row in range(2, 8)}
    labels = [formula_sheet[f"A{row}"].value for row in range(2, 8)]
    normalized_formulas = {cell: V1.formula_norm(value) for cell, value in formula_values.items()}
    formula_findings = [
        (labels == V1.METRIC_ORDER, "Metrics 指标名称或顺序错误"),
        (str(formula_values["B2"] or "").startswith("=") and all(token in normalized_formulas["B2"] for token in ("COUNTA", "DATA_ANALYSIS!A2:A5")), "B2 不是从 Data_Analysis 推导的总论文数公式"),
        (not str(formula_values["B3"] or "").startswith("=") and V1.close(formula_values["B3"], 1), "B3 排除数应为常量 1"),
        (str(formula_values["B4"] or "").startswith("=") and all(token in normalized_formulas["B4"] for token in ("COUNT", "DATA_ANALYSIS!F2:F5")), "B4 不是从 Growth_Pct 推导的匹配数公式"),
        (str(formula_values["B5"] or "").startswith("=") and all(token in normalized_formulas["B5"] for token in ("AVERAGE", "DATA_ANALYSIS!F2:F5")), "B5 不是从 Growth_Pct 推导的平均值公式"),
        (str(formula_values["B6"] or "").startswith("=") and all(token in normalized_formulas["B6"] for token in ("INDEX", "MAX", "DATA_ANALYSIS!E2:E5", "DATA_ANALYSIS!F2:F5")), "B6 不是最高增长领域公式"),
        (str(formula_values["B7"] or "").startswith("=") and all(token in normalized_formulas["B7"] for token in ("COUNT", "COUNTA", "DATA_ANALYSIS!F2:F5", "DATA_ANALYSIS!A2:A5")), "B7 不是覆盖率公式"),
    ]
    formula_book.close()
    # 公式结果由 Data_Analysis 真值与精确依赖范围共同决定；真实 LibreOffice 重算及可见值另由 R013 独立验证。
    findings = formula_findings
    score, failures, earned, total = _ratio(findings)
    return score, failures, earned, total, {V1.METRIC_ORDER[index]: formula_values[f"B{index + 2}"] for index in range(6)}


def _recommendations(artifact: Path, delivered: dict[str, Any], expected: dict[str, Any]) -> tuple[float, list[str], int, int]:
    book = load_workbook(artifact / REPORT, data_only=True)
    rows, missing = V1.read_sheet_rows(book["Recommendations"], V1.REC_HEADERS)
    book.close()
    findings = [(not missing, f"Recommendations 缺表头：{missing}")]
    findings.extend(_recommendation_findings(rows, expected["recommendations"], "Report.Recommendations"))
    findings.extend(_recommendation_findings(delivered.get("recommendations"), expected["recommendations"], "JSON.recommendations"))
    for index in range(3):
        excel_row = rows[index] if index < len(rows) else {}
        json_rows = delivered.get("recommendations") if isinstance(delivered.get("recommendations"), list) else []
        json_row = json_rows[index] if index < len(json_rows) and isinstance(json_rows[index], dict) else {}
        findings.append((all(_same(excel_row.get(key), json_row.get(key)) for key in ("Rank", "Priority", "Area", "Growth_Pct")), f"第 {index + 1} 条建议跨 JSON/Excel 冲突"))
    return _ratio(findings)


def _tracker(artifact: Path, expected: dict[str, Any]) -> tuple[float, list[str], int, int, dict[str, Any]]:
    expected_rows = [{
        "Paper_ID": row["Paper_ID"],
        "Title": row["Title"],
        "Primary_Category": row["Primary_Category"],
        "Benchmark_Area": row["Benchmark_Area"],
        "Growth_Pct": row["Growth_Pct"],
        "Review_Status": "Ready for review" if row["Match_Status"] == "Matched" else "Needs separate benchmark",
    } for row in expected["data_analysis"]]
    book = load_workbook(artifact / TRACKER, data_only=True)
    rows, missing = V1.read_sheet_rows(book["Arxiv Litreview Tracker"], V1.TRACKER_HEADERS)
    book.close()
    findings = [(not missing, f"Tracker 缺表头：{missing}")]
    findings.extend(_rows_findings(rows, expected_rows, V1.TRACKER_HEADERS, "Tracker"))
    score, failures, earned, total = _ratio(findings)
    return score, failures, earned, total, {str(row.get("Paper_ID")): row for row in rows}


def _word(artifact: Path, expected: dict[str, Any]) -> tuple[float, list[str], int, int, str]:
    text, headings, table_count, inline_shapes = V1.docx_text_and_structure(artifact / WORD)
    normalized = V1.normalized(text)
    required_headings = ["краткое резюме", "источники и метод", "ключевые выводы", "рекомендации", "ограничения"]
    heading_norm = [V1.normalized(value) for value in headings]
    tokens = [
        "arxiv_papers.json", "benchmark_data.json", "5", "4", "3", "21,67%", "75,00%", "llms", "45%",
        "computer vision", "12%", "rl", "8%", "quant-ph", "federated learning",
        *[row["Paper_ID"] for row in expected["data_analysis"]],
    ]
    findings = [
        (len(text) >= 1800, "Word 正文少于 1800 字符"),
        (len(re.findall(r"[А-Яа-яЁё]", text)) >= 1000, "Word 俄文字符少于 1000"),
        (table_count >= 2, "Word 原生表格少于 2 张"),
        (inline_shapes == 0, "Word 含图片替代关键内容"),
    ]
    findings.extend((heading in heading_norm, f"Word 缺标题 {heading}") for heading in required_headings)
    findings.extend((V1.normalized(token) in normalized, f"Word 缺事实 {token}") for token in tokens)
    score, failures, earned, total = _ratio(findings)
    return score, failures, earned, total, text


def _method_boundaries(delivered: dict[str, Any], word_text: str) -> tuple[float, list[str], int, int]:
    json_blob = V1.normalized(json.dumps(delivered, ensure_ascii=False))
    word_blob = V1.normalized(word_text)
    findings = []
    for token in ("arxiv_papers.json", "benchmark_data.json"):
        findings.extend([(token in json_blob, f"JSON 未声明来源 {token}"), (token in word_blob, f"Word 未声明来源 {token}")])
    for token in ("primary_category", "language model", "vision transformer", "reinforcement learning"):
        findings.append((token in word_blob, f"Word 未声明匹配方法 {token}"))
    limitations = delivered.get("limitations")
    findings.extend([
        (isinstance(limitations, list) and len(limitations) >= 4, "JSON 限制边界少于 4 条"),
        ("нет счётчиков цитирования" in word_blob, "Word 未说明没有引用数"),
        ("gap-анализ" in word_blob, "Word 未说明不能做 gap 分析"),
        ("нет временного ряда" in word_blob, "Word 未说明没有时间序列"),
        ("не измерение качества" in word_blob, "Word 未说明增长率不是论文质量"),
    ])
    return _ratio(findings)


def _cross(delivered: dict[str, Any], report_rows: dict[str, Any], tracker_rows: dict[str, Any], word_text: str) -> tuple[float, list[str], int, int]:
    json_rows = {str(row.get("Paper_ID")): row for row in delivered.get("data_analysis", []) if isinstance(row, dict)}
    findings: list[tuple[bool, str]] = [(set(json_rows) == set(report_rows) == set(tracker_rows), "JSON/Report/Tracker 论文集合冲突")]
    for paper_id in sorted(set(json_rows) | set(report_rows) | set(tracker_rows)):
        json_row, report_row, tracker_row = json_rows.get(paper_id, {}), report_rows.get(paper_id, {}), tracker_rows.get(paper_id, {})
        for key in ("Title", "Primary_Category", "Benchmark_Area", "Growth_Pct"):
            findings.append((_same(json_row.get(key), report_row.get(key)) and _same(json_row.get(key), tracker_row.get(key)), f"{paper_id} 的 {key} 跨文件冲突"))
        findings.append((paper_id in word_text, f"Word 未覆盖 {paper_id}"))
    return _ratio(findings)


def _formatting(artifact: Path) -> tuple[float, list[str], int, int]:
    report = load_workbook(artifact / REPORT, data_only=False)
    tracker = load_workbook(artifact / TRACKER, data_only=False)
    findings: list[tuple[bool, str]] = []
    for label, book, sheets in (("Report", report, ["Data_Analysis", "Metrics", "Recommendations"]), ("Tracker", tracker, ["Arxiv Litreview Tracker"])):
        for sheet_name in sheets:
            sheet = book[sheet_name]
            headers = list(sheet[1])
            findings.extend([
                (len(sheet.tables) >= 1, f"{label}.{sheet_name} 没有原生 Excel 表"),
                (str(sheet.freeze_panes or "") == "A2", f"{label}.{sheet_name} 未冻结首行"),
                (all(bool(cell.font.bold) for cell in headers), f"{label}.{sheet_name} 表头未全部加粗"),
                (all(cell.fill.fill_type == "solid" for cell in headers), f"{label}.{sheet_name} 表头未全部填色"),
                (all((sheet.column_dimensions[cell.column_letter].width or 0) >= 9 for cell in headers), f"{label}.{sheet_name} 存在过窄列"),
            ])
    report.close()
    tracker.close()
    document = Document(artifact / WORD)
    findings.extend([(len(document.tables) >= 2, "Word 原生表格少于 2 张"), (sum(paragraph.style.name.startswith("Heading") for paragraph in document.paragraphs) >= 5, "Word 标题层级少于 5 个")])
    return _ratio(findings)


def _render(artifact: Path) -> tuple[float, list[str], int, int]:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-arxiv-render-") as raw:
        recalculated, texts, pages, evidence = V1.libreoffice_prepare(artifact / REPORT, artifact / TRACKER, artifact / WORD, Path(raw))
    report_text = V1.normalized(texts.get(REPORT, ""))
    tracker_text = V1.normalized(texts.get(TRACKER, ""))
    word_text = V1.normalized(texts.get(WORD, ""))
    findings = [
        (recalculated is not None, f"Report 未重算：{evidence}"),
        (pages.get(REPORT, 0) >= 3, "Report 渲染页数少于 3"),
        (pages.get(TRACKER, 0) >= 1, "Tracker 渲染失败"),
        (pages.get(WORD, 0) >= 2, "Word 渲染页数少于 2"),
        (all(token in report_text for token in ("paper_id", "2301.01234", "total_papers", "rank", "priority", "action")), "Report 渲染缺核心文字"),
        (all(token in tracker_text for token in ("paper_id", "2304.03456", "review_status")), "Tracker 渲染缺核心文字"),
        (all(token in word_text for token in ("краткое резюме", "llms", "45%", "ограничения")), "Word 渲染缺核心文字"),
    ]
    return _ratio(findings)


def _boundary_penalty(delivered: dict[str, Any], report_rows: dict[str, Any], tracker_rows: dict[str, Any]) -> tuple[float, list[str], float]:
    relevant_sets = []
    for rows in (delivered.get("data_analysis"), list(report_rows.values()), list(tracker_rows.values())):
        relevant_sets.append({str(row.get("Paper_ID")) for row in rows or [] if isinstance(row, dict)})
    contaminated = any("2306.05678" in values for values in relevant_sets)
    return (1.0 if contaminated else 0.0), (["quant-ph 噪声论文被纳入主分析"] if contaminated else []), 1.0


def _script_penalty(canonical_score: float, mutated_score: float, canonical: dict[str, Any] | None, mutated: dict[str, Any] | None) -> tuple[float, list[str], float]:
    labels: list[str] = []
    severity = 0.0
    if canonical is None:
        severity = max(severity, 0.65)
        labels.append("处理脚本在原始输入上无法运行")
    if mutated is None:
        severity = max(severity, 0.8)
        labels.append("处理脚本在变异输入上无法运行")
    elif canonical is not None and canonical == mutated:
        severity = 1.0
        labels.append("处理脚本对 Paper_ID 和增长率变异完全无响应")
    elif mutated_score < 0.5 and canonical_score > 0.9:
        severity = max(severity, 0.75)
        labels.append("处理脚本疑似写死原始结果")
    return severity, labels, 1.0


def _extra_penalty(artifact: Path) -> tuple[float, list[str], float]:
    # solution/ 中的回放入口不是 Agent 交付；Golden 自检时允许存在，真实 /app/output 不会带入。
    # 执行可回放 Python 交付自然产生的 __pycache__ 也不是额外交付或信息泄漏。
    allowed = {*REQUIRED, "solve.py", "solve.sh", "__pycache__"}
    extras = sorted(path.name for path in artifact.iterdir() if path.name not in allowed)
    return (1.0 if extras else 0.0), ([f"输出目录包含未要求的文件或目录：{extras}"] if extras else []), 1.0


def _unsupported_penalty(delivered: dict[str, Any], artifact: Path, word_text: str) -> tuple[float, list[str], float]:
    report = load_workbook(artifact / REPORT, data_only=True)
    tracker = load_workbook(artifact / TRACKER, data_only=True)
    report_headers = [str(cell.value or "") for sheet in report.worksheets for cell in sheet[1]]
    tracker_headers = [str(cell.value or "") for cell in tracker["Arxiv Litreview Tracker"][1]]
    office_values = "\n".join(str(cell.value or "") for book in (report, tracker) for sheet in book.worksheets for row in sheet.iter_rows() for cell in row)
    report.close()
    tracker.close()
    findings = V1.unsupported_assertions(delivered, report_headers, tracker_headers, f"{office_values}\n{word_text}")
    severity = min(1.0, 0.5 + 0.15 * len(findings)) if findings else 0.0
    return severity, findings, 1.0


def _safe_positive(call, *args):
    try:
        return call(*args)
    except Exception as exc:
        return 0.0, [f"任务专属检查异常：{type(exc).__name__}: {exc}"], 0, 1


def _safe_penalty(call, *args):
    try:
        return call(*args)
    except Exception as exc:
        return 1.0, [f"严重边界无法核验：{type(exc).__name__}: {exc}"], 1.0


def evaluate(artifact: Path) -> dict[str, Any]:
    ok, errors = _core(artifact)
    hashes = [_sha256(artifact / name) for name in REQUIRED if (artifact / name).is_file()]
    if not ok:
        return _fatal(errors, hashes)

    papers_path, benchmark_path = _task_inputs()
    expected = V1.derive_expected(papers_path, benchmark_path)
    delivered = V1.load_result(artifact / RESULT_JSON)
    results = [_criterion("R001", 1.0, "五份核心交付存在、非空、可解析且工作表结构正确", 5, 5)]
    scores: dict[str, float] = {"R001": 1.0}

    canonical_score, canonical_failures, canonical_earned, canonical_total, canonical = _processor_checks(artifact, False)
    mutated_score, mutated_failures, mutated_earned, mutated_total, mutated = _processor_checks(artifact, True)
    for key, score, failures, earned, total in (
        ("R002", canonical_score, canonical_failures, canonical_earned, canonical_total),
        ("R003", mutated_score, mutated_failures, mutated_earned, mutated_total),
    ):
        results.append(_criterion(key, score, "全部任务专属原子检查通过" if not failures else "未通过：" + "；".join(failures), earned, total))
        scores[key] = score

    json_score, json_failures, json_earned, json_total = _ratio(_result_findings(delivered, expected))
    results.append(_criterion("R004", json_score, "JSON 全部原子事实正确" if not json_failures else "未通过：" + "；".join(json_failures), json_earned, json_total))
    scores["R004"] = json_score

    report_score, report_failures, report_earned, report_total, report_rows = _safe_positive(_report_rows, artifact, expected)
    metric_score, metric_failures, metric_earned, metric_total, _ = _metrics(artifact, expected)
    rec_score, rec_failures, rec_earned, rec_total = _safe_positive(_recommendations, artifact, delivered, expected)
    tracker_score, tracker_failures, tracker_earned, tracker_total, tracker_rows = _safe_positive(_tracker, artifact, expected)
    word_score, word_failures, word_earned, word_total, word_text = _word(artifact, expected)
    method_score, method_failures, method_earned, method_total = _safe_positive(_method_boundaries, delivered, word_text)
    cross_score, cross_failures, cross_earned, cross_total = _safe_positive(_cross, delivered, report_rows, tracker_rows, word_text)
    format_score, format_failures, format_earned, format_total = _safe_positive(_formatting, artifact)
    render_score, render_failures, render_earned, render_total = _safe_positive(_render, artifact)
    checks = {
        "R005": (report_score, report_failures, report_earned, report_total),
        "R006": (metric_score, metric_failures, metric_earned, metric_total),
        "R007": (rec_score, rec_failures, rec_earned, rec_total),
        "R008": (tracker_score, tracker_failures, tracker_earned, tracker_total),
        "R009": (word_score, word_failures, word_earned, word_total),
        "R010": (method_score, method_failures, method_earned, method_total),
        "R011": (cross_score, cross_failures, cross_earned, cross_total),
        "R012": (format_score, format_failures, format_earned, format_total),
        "R013": (render_score, render_failures, render_earned, render_total),
    }
    for key, (score, failures, earned, total) in checks.items():
        results.append(_criterion(key, score, "全部任务专属原子检查通过" if not failures else "未通过：" + "；".join(failures), earned, total))
        scores[key] = score

    penalties = {
        "R014": _safe_penalty(_boundary_penalty, delivered, report_rows, tracker_rows),
        "R015": _safe_penalty(_script_penalty, canonical_score, mutated_score, canonical, mutated),
        "R016": _safe_penalty(_extra_penalty, artifact),
        "R017": _safe_penalty(_unsupported_penalty, delivered, artifact, word_text),
    }
    for key, (score, labels, total) in penalties.items():
        results.append(_criterion(key, score, "未触发任务专属严重边界" if not labels else "；".join(labels), score * total, total))
        scores[key] = score

    total_weight = sum(POSITIVE_WEIGHTS.values())
    positive = sum(POSITIVE_WEIGHTS[key] * scores[key] for key in POSITIVE_WEIGHTS)
    penalty = sum(abs(PENALTY_WEIGHTS[key]) * scores[key] for key in PENALTY_WEIGHTS)
    reward = round(max(0.0, min(1.0, (positive - penalty) / total_weight)), 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": hashes,
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= 0.8,
        "hurdle_failures": [],
        "penalties": [{"id": key, "amount": round(abs(PENALTY_WEIGHTS[key]) * scores[key] / total_weight, 6)} for key in PENALTY_WEIGHTS if scores[key] > 0],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    parser.add_argument("--scoring-version", choices=("v1", "v2"), default="v2")
    args = parser.parse_args()
    try:
        result = V1.evaluate(args.artifact_dir) if args.scoring_version == "v1" else evaluate(args.artifact_dir)
    except Exception as exc:
        result = _fatal([f"Judge 异常：{type(exc).__name__}: {exc}"], [])
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(rendered + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(rendered)


if __name__ == "__main__":
    main()
