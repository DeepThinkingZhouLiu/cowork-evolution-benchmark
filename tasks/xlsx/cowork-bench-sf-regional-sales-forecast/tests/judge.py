#!/usr/bin/env python3
"""Cowork-Bench 区域销售预测题的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


XLSX_NAME = "Sales_Forecast.xlsx"
DOCX_NAME = "Forecast_Report.docx"
TASK_ID = "cowork-evo/cowork-bench-sf-regional-sales-forecast"
RUBRIC_VERSION = "cowork-bench-sf-regional-sales-static-v0.1"
JUDGE_VERSION = "cowork-bench-sf-regional-sales-static-v0.1"
HURDLE_CAPS = {identifier: 0.74 for identifier in ("R002", "R003", "R004", "R005", "R006", "R008")}
HURDLE_CAPS["R001"] = 0.0

REGIONS = [
    "Азиатско-Тихоокеанский регион",
    "Европа",
    "Латинская Америка",
    "Ближний Восток",
    "Северная Америка",
]
HISTORICAL_HEADERS = ["Region", "Month", "Orders", "Revenue"]
FORECAST_HEADERS = [
    "Region", "Trailing_3M_Avg", "GDP_Growth_Pct", "Economic_Adjustment",
    "Adjusted_Monthly_Forecast", "Next_Quarter_Forecast",
]
SUMMARY_HEADERS = ["Metric", "Value"]

CRITERION_SPECS = {
    "R001": ("指定 Excel 和 Word 可打开、可编辑且结构完整", 4, "hurdle"),
    "R002": ("Historical 的 15 条地区月度来源记录全部正确", 6, "hurdle"),
    "R003": ("五个地区的三月移动平均全部正确", 5, "hurdle"),
    "R004": ("五个地区的 GDP 联结和经济调整系数全部正确", 5, "hurdle"),
    "R005": ("五个地区的调整月预测和下季度预测全部正确", 6, "hurdle"),
    "R006": ("Summary 的十项总额、增长率和区域指标全部正确", 6, "hurdle"),
    "R007": ("Word 报告含五个规定的俄文章节", 4, "positive"),
    "R008": ("Word 报告的总额、五地区数字和建议与 Excel 一致", 5, "hurdle"),
    "R009": ("Excel 和 Word 均可由 LibreOffice 真实渲染且关键文本可读", 3, "positive"),
    "R010": ("任一历史记录、地区联结或核心预测计算错误", -10, "penalty"),
    "R011": ("Word 缺章节、缺关键数字或与 Excel 冲突", -10, "penalty"),
    "R012": ("任一交付物损坏、截图化或无法渲染", -10, "penalty"),
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


def normalized(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def to_float(value) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def close(actual, expected, tolerance=0.011) -> bool:
    number = to_float(actual)
    return number is not None and abs(number - float(expected)) <= tolerance


def derive_expected(input_dir: Path):
    historical = json.loads((input_dir / "historical_sales.json").read_text(encoding="utf-8"))
    economics = json.loads((input_dir / "economic_indicators.json").read_text(encoding="utf-8"))["regions"]
    historical_lookup = {(row["region"], row["month"]): row for row in historical}
    revenues = defaultdict(list)
    for row in historical:
        revenues[row["region"]].append(float(row["revenue"]))
    forecasts = {}
    for region in REGIONS:
        raw_average = sum(revenues[region]) / 3
        trailing = round(raw_average, 2)
        gdp = float(economics[region]["gdp_growth"])
        adjustment = round(1 + gdp / 100, 4)
        monthly = round(trailing * adjustment, 2)
        quarterly = round(monthly * 3, 2)
        forecasts[region] = {
            "trailing": trailing,
            "gdp": gdp,
            "adjustment": adjustment,
            "monthly": monthly,
            "quarterly": quarterly,
            "unrounded_quarterly": raw_average * adjustment * 3,
        }
    total_q4 = round(sum(float(row["revenue"]) for row in historical), 2)
    total_q1 = round(sum(round(item["unrounded_quarterly"], 2) for item in forecasts.values()), 2)
    highest = max(forecasts, key=lambda region: forecasts[region]["unrounded_quarterly"])
    lowest = min(forecasts, key=lambda region: forecasts[region]["unrounded_quarterly"])
    highest_gdp = max(forecasts, key=lambda region: forecasts[region]["gdp"])
    summary = {
        "Total_Q4_2025_Revenue": total_q4,
        "Total_Monthly_Average": round(total_q4 / 3, 2),
        "Total_Q1_2026_Forecast": total_q1,
        "QoQ_Growth_Pct": round((total_q1 / total_q4 - 1) * 100, 1),
        "Highest_Forecast_Region": highest,
        "Highest_Forecast_Amount": round(forecasts[highest]["unrounded_quarterly"], 2),
        "Lowest_Forecast_Region": lowest,
        "Lowest_Forecast_Amount": round(forecasts[lowest]["unrounded_quarterly"], 2),
        "Highest_GDP_Growth_Region": highest_gdp,
        "Regions_Count": 5,
    }
    return historical_lookup, forecasts, summary


def sheet_records(workbook, sheet_name: str, headers: list[str]) -> tuple[list[dict], list[str]]:
    if sheet_name not in workbook.sheetnames:
        return [], headers
    rows = list(workbook[sheet_name].iter_rows(values_only=True))
    if not rows:
        return [], headers
    row_headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    index = {name: position for position, name in enumerate(row_headers)}
    missing = [name for name in headers if name not in index]
    if missing:
        return [], missing
    records = []
    for row in rows[1:]:
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        records.append({name: row[index[name]] if index[name] < len(row) else None for name in headers})
    return records, []


def libreoffice_prepare(xlsx_path: Path, docx_path: Path, root: Path) -> tuple[Path | None, dict[str, str], str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return None, {}, "缺少 libreoffice 或 pdftotext"
    source_dir = root / "source"
    recalc_dir = root / "recalc"
    pdf_dir = root / "pdf"
    source_dir.mkdir()
    recalc_dir.mkdir()
    pdf_dir.mkdir()
    xlsx_copy = source_dir / XLSX_NAME
    docx_copy = source_dir / DOCX_NAME
    shutil.copy2(xlsx_path, xlsx_copy)
    shutil.copy2(docx_path, docx_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()

    recalc = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(xlsx_copy)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    recalculated = recalc_dir / XLSX_NAME
    if recalc.returncode != 0 or not recalculated.is_file():
        return None, {}, (recalc.stderr or recalc.stdout or "Excel 重算失败")[-1000:]

    texts = {}
    for source in (recalculated, docx_copy):
        converted = subprocess.run(
            [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(source)],
            capture_output=True, text=True, timeout=300, check=False,
        )
        pdf_path = pdf_dir / f"{source.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return recalculated, texts, (converted.stderr or converted.stdout or f"{source.name} PDF 失败")[-1000:]
        extracted = subprocess.run([pdftotext, "-raw", str(pdf_path), "-"], capture_output=True, text=True, timeout=120, check=False)
        if extracted.returncode != 0:
            return recalculated, texts, (extracted.stderr or f"{source.name} PDF 文本失败")[-1000:]
        texts[source.suffix.lower()] = extracted.stdout
    return recalculated, texts, "Excel 重算成功，Excel/Word PDF 渲染成功"


def evaluate(artifact_dir: Path) -> dict:
    xlsx_path = artifact_dir / XLSX_NAME
    docx_path = artifact_dir / DOCX_NAME
    results = []
    if not xlsx_path.is_file() or not docx_path.is_file():
        results.append(criterion("R001", False, f"Excel 存在={xlsx_path.is_file()}，Word 存在={docx_path.is_file()}"))
        return build_result(results, [xlsx_path, docx_path])

    input_dir = Path(__file__).resolve().parents[1] / "data" / "input_files"
    expected_historical, expected_forecasts, expected_summary = derive_expected(input_dir)

    with tempfile.TemporaryDirectory(prefix="cowork-evo-sales-forecast-") as temp_dir:
        recalculated, rendered_texts, render_evidence = libreoffice_prepare(xlsx_path, docx_path, Path(temp_dir))
        workbook_path = recalculated or xlsx_path
        try:
            if not zipfile.is_zipfile(workbook_path) or not zipfile.is_zipfile(docx_path):
                raise ValueError("交付物不是有效 OOXML ZIP")
            workbook = load_workbook(workbook_path, data_only=True)
            historical_rows, historical_missing = sheet_records(workbook, "Historical", HISTORICAL_HEADERS)
            forecast_rows, forecast_missing = sheet_records(workbook, "Forecast", FORECAST_HEADERS)
            summary_rows, summary_missing = sheet_records(workbook, "Summary", SUMMARY_HEADERS)
            sheet_names = workbook.sheetnames
            workbook.close()
            document = Document(docx_path)
            paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
            word_text = "\n".join(paragraphs)
            media_count = 0
            with zipfile.ZipFile(docx_path) as archive:
                media_count = sum(1 for name in archive.namelist() if name.startswith("word/media/") and not name.endswith("/"))
        except Exception as error:
            results.append(criterion("R001", False, f"无法读取交付物：{error}"))
            return build_result(results, [xlsx_path, docx_path])

    structure_ok = (
        sheet_names == ["Historical", "Forecast", "Summary"]
        and not historical_missing and not forecast_missing and not summary_missing
        and len(paragraphs) >= 10 and len(normalized(word_text)) >= 1200 and media_count == 0
    )
    results.append(
        criterion(
            "R001", structure_ok,
            f"工作表={sheet_names}，缺失表头={historical_missing + forecast_missing + summary_missing}，Word段落={len(paragraphs)}，Word字符={len(normalized(word_text))}，媒体={media_count}",
        )
    )

    actual_historical = {}
    historical_duplicates = []
    for row in historical_rows:
        key = (str(row["Region"] or "").strip(), str(row["Month"] or "").strip())
        if key in actual_historical:
            historical_duplicates.append(key)
        actual_historical[key] = row
    missing_historical = sorted(set(expected_historical) - set(actual_historical))
    extra_historical = sorted(set(actual_historical) - set(expected_historical))
    historical_errors = []
    for key in set(expected_historical) & set(actual_historical):
        actual = actual_historical[key]
        expected = expected_historical[key]
        if not close(actual["Orders"], expected["orders"], 0.001) or not close(actual["Revenue"], expected["revenue"], 0.011):
            historical_errors.append((key, actual["Orders"], actual["Revenue"], expected["orders"], expected["revenue"]))
    historical_ok = len(historical_rows) == 15 and not missing_historical and not extra_historical and not historical_duplicates and not historical_errors
    results.append(criterion("R002", historical_ok, f"行数={len(historical_rows)}/15，缺失={missing_historical[:4]}，多余={extra_historical[:4]}，重复={historical_duplicates[:4]}，数值错误={historical_errors[:4]}"))

    actual_forecasts = {str(row["Region"] or "").strip(): row for row in forecast_rows}
    region_set_ok = len(forecast_rows) == 5 and set(actual_forecasts) == set(REGIONS)
    trailing_errors = []
    adjustment_errors = []
    forecast_errors = []
    for region in set(REGIONS) & set(actual_forecasts):
        actual = actual_forecasts[region]
        expected = expected_forecasts[region]
        if not close(actual["Trailing_3M_Avg"], expected["trailing"]):
            trailing_errors.append((region, actual["Trailing_3M_Avg"], expected["trailing"]))
        if not close(actual["GDP_Growth_Pct"], expected["gdp"], 0.0001) or not close(actual["Economic_Adjustment"], expected["adjustment"], 0.00011):
            adjustment_errors.append((region, actual["GDP_Growth_Pct"], actual["Economic_Adjustment"], expected["gdp"], expected["adjustment"]))
        if not close(actual["Adjusted_Monthly_Forecast"], expected["monthly"]) or not close(actual["Next_Quarter_Forecast"], expected["quarterly"]):
            forecast_errors.append((region, actual["Adjusted_Monthly_Forecast"], actual["Next_Quarter_Forecast"], expected["monthly"], expected["quarterly"]))
    trailing_ok = region_set_ok and not trailing_errors
    adjustment_ok = region_set_ok and not adjustment_errors
    forecast_ok = region_set_ok and not forecast_errors
    results.append(criterion("R003", trailing_ok, f"地区集合={region_set_ok}，移动平均错误={trailing_errors}"))
    results.append(criterion("R004", adjustment_ok, f"地区集合={region_set_ok}，GDP/调整错误={adjustment_errors}"))
    results.append(criterion("R005", forecast_ok, f"地区集合={region_set_ok}，预测错误={forecast_errors}"))

    actual_summary = {str(row["Metric"] or "").strip(): row["Value"] for row in summary_rows}
    summary_errors = []
    for metric, expected in expected_summary.items():
        if metric not in actual_summary:
            summary_errors.append((metric, "缺失", expected))
        elif isinstance(expected, str):
            if normalized(actual_summary[metric]) != normalized(expected):
                summary_errors.append((metric, actual_summary[metric], expected))
        else:
            tolerance = 0.11 if metric == "QoQ_Growth_Pct" else 0.011
            if not close(actual_summary[metric], expected, tolerance):
                summary_errors.append((metric, actual_summary[metric], expected))
    summary_ok = len(summary_rows) == 10 and len(actual_summary) == 10 and not summary_errors
    results.append(criterion("R006", summary_ok, f"指标={len(actual_summary)}/10，错误={summary_errors[:6]}"))

    text_norm = normalized(word_text)
    section_checks = {
        "summary": "краткое резюме" in text_norm,
        "methodology": "методолог" in text_norm,
        "regional": "региональн" in text_norm and "анализ" in text_norm,
        "assumptions": "допущен" in text_norm,
        "recommendations": "рекомендац" in text_norm,
    }
    report_structure_ok = all(section_checks.values()) and len(paragraphs) >= 10
    results.append(criterion("R007", report_structure_ok, f"章节={section_checks}，段落={len(paragraphs)}"))

    digits = re.sub(r"\D", "", word_text)
    required_number_tokens = ["393653", "402638", "93992", "73711", "76586", "75943", "82404"]
    missing_numbers = [token for token in required_number_tokens if token not in digits]
    missing_regions = [region for region in REGIONS if normalized(region) not in text_norm]
    formula_ok = all(token in text_norm for token in ("скользящ", "ввп", "100"))
    recommendation_ok = "азиатско-тихоокеанск" in text_norm and "европ" in text_norm
    report_content_ok = not missing_numbers and not missing_regions and formula_ok and recommendation_ok
    results.append(criterion("R008", report_content_ok, f"缺数字={missing_numbers}，缺地区={missing_regions}，公式={formula_ok}，建议={recommendation_ok}"))

    xlsx_rendered = normalized(rendered_texts.get(".xlsx", ""))
    docx_rendered = normalized(rendered_texts.get(".docx", ""))
    xlsx_render_ok = all(token in xlsx_rendered for token in ("region", "2025-10", "93992", "402638"))
    docx_render_ok = all(token in docx_rendered for token in ("методолог", "рекомендац", "402"))
    render_ok = recalculated is not None and xlsx_render_ok and docx_render_ok
    results.append(criterion("R009", render_ok, f"{render_evidence}，Excel文本={xlsx_render_ok}，Word文本={docx_render_ok}"))

    model_error = not (historical_ok and trailing_ok and adjustment_ok and forecast_ok and summary_ok)
    report_error = not (report_structure_ok and report_content_ok)
    broken = not structure_ok or not render_ok
    results.append(criterion("R010", model_error, f"历史={historical_ok}，平均={trailing_ok}，调整={adjustment_ok}，预测={forecast_ok}，摘要={summary_ok}"))
    results.append(criterion("R011", report_error, f"报告结构={report_structure_ok}，报告数值={report_content_ok}"))
    results.append(criterion("R012", broken, f"结构={structure_ok}，渲染={render_ok}，媒体={media_count}"))
    return build_result(results, [xlsx_path, docx_path])


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
        criterion_results.append({
            "criterion_id": item["criterion_id"], "score": score, "evidence": item["evidence"], "method": item["method"],
            "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]},
        })
    return {
        "task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests, "criterion_results": criterion_results,
        "reward": round(reward, 6), "passed": reward >= 0.75, "hurdle_failures": hurdle_failures,
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
