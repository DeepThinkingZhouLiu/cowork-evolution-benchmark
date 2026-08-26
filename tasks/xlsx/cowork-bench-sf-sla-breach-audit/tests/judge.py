#!/usr/bin/env python3
"""从静态工单与 SLA 策略重算真值的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from email import policy
from email.parser import Parser
from pathlib import Path

from openpyxl import load_workbook


WORKBOOK_NAME = "SLA_Breach_Audit.xlsx"
EMAIL_NAME = "SLA_Breach_Audit_Email.txt"
TASK_ID = "cowork-evo/cowork-bench-sf-sla-breach-audit"
RUBRIC_VERSION = "cowork-bench-sf-sla-breach-audit-static-v0.1"
JUDGE_VERSION = "cowork-bench-sf-sla-breach-audit-static-v0.1"
PASS_THRESHOLD = 0.8
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.45,
    "R003": 0.5,
    "R004": 0.4,
    "R005": 0.4,
    "R006": 0.35,
    "R007": 0.79,
    "R009": 0.7,
}

ANALYSIS_HEADERS = [
    "Priority",
    "Total_Tickets",
    "Avg_Response_Hours",
    "SLA_Target_Hours",
    "Breached_Count",
    "Breach_Rate",
]
SUMMARY_METRICS = [
    "Total_Tickets",
    "Total_Breached",
    "Overall_Breach_Rate",
    "Worst_Priority",
    "Best_Priority",
]

CRITERION_SPECS = {
    "R001": ("两个指定文件可打开且 Excel 恰有两张指定表", 5, "hurdle"),
    "R002": ("表头、优先级集合和 SLA 目标排序正确", 5, "hurdle"),
    "R003": ("工单数、平均响应时间和 SLA 目标正确", 7, "hurdle"),
    "R004": ("三个优先级的违约数精确正确", 8, "hurdle"),
    "R005": ("三个优先级的违约率正确", 7, "hurdle"),
    "R006": ("Summary 五项指标的名称、顺序和值正确", 8, "hurdle"),
    "R007": ("关键比率和汇总值使用可审计的 Excel 公式", 4, "hurdle"),
    "R008": ("工作簿格式、冻结、筛选和列宽达到专业交付要求", 4, "positive"),
    "R009": ("邮件草稿的头部与核心结论正确且跨文件一致", 5, "hurdle"),
    "R010": ("LibreOffice 打开、重算、另存和渲染通过", 3, "positive"),
    "R011": ("核心审计数据存在错误", -12, "penalty"),
    "R012": ("交付物缺失、损坏或邮件与 Excel 冲突", -10, "penalty"),
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


def round_2(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def normalized(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def close(actual, expected, tolerance: float = 0.001) -> bool:
    try:
        return abs(float(actual) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return False


def derive_expected(input_dir: Path) -> tuple[list[dict], dict]:
    with (input_dir / "sla_policies.csv").open(encoding="utf-8", newline="") as handle:
        policies = {
            row["PRIORITY"]: Decimal(row["RESPONSE_TARGET_HOURS"])
            for row in csv.DictReader(handle)
        }

    aggregates = defaultdict(lambda: {"count": 0, "sum": Decimal("0"), "breached": 0})
    with (input_dir / "support_tickets.csv").open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            priority = row["PRIORITY"]
            if priority not in policies:
                raise ValueError(f"工单优先级没有 SLA 策略：{priority}")
            response = Decimal(row["RESPONSE_TIME_HOURS"])
            aggregates[priority]["count"] += 1
            aggregates[priority]["sum"] += response
            aggregates[priority]["breached"] += response > policies[priority]

    rows = []
    for priority in sorted(aggregates, key=lambda item: (policies[item], item)):
        item = aggregates[priority]
        rows.append(
            {
                "Priority": priority,
                "Total_Tickets": item["count"],
                "Avg_Response_Hours": round_2(item["sum"] / item["count"]),
                "SLA_Target_Hours": policies[priority],
                "Breached_Count": item["breached"],
                "Breach_Rate": round_2(Decimal(item["breached"]) * 100 / item["count"]),
            }
        )
    total = sum(row["Total_Tickets"] for row in rows)
    breached = sum(row["Breached_Count"] for row in rows)
    summary = {
        "Total_Tickets": total,
        "Total_Breached": breached,
        "Overall_Breach_Rate": round_2(Decimal(breached) * 100 / total),
        "Worst_Priority": max(rows, key=lambda row: row["Breach_Rate"])["Priority"],
        "Best_Priority": min(rows, key=lambda row: row["Breach_Rate"])["Priority"],
    }
    return rows, summary


def libreoffice_recalculate(workbook_path: Path, root: Path) -> tuple[Path | None, str, str]:
    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if libreoffice is None or pdftotext is None:
        return None, "", "缺少 libreoffice 或 pdftotext"
    source_dir = root / "source"
    recalc_dir = root / "recalc"
    pdf_dir = root / "pdf"
    source_dir.mkdir()
    recalc_dir.mkdir()
    pdf_dir.mkdir()
    source_copy = source_dir / WORKBOOK_NAME
    shutil.copy2(workbook_path, source_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()
    recalc = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(source_copy)],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    recalculated = recalc_dir / WORKBOOK_NAME
    if recalc.returncode != 0 or not recalculated.is_file():
        return None, "", (recalc.stderr or recalc.stdout or "LibreOffice 重算失败")[-1200:]
    render = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(recalculated)],
        capture_output=True,
        text=True,
        timeout=300,
        check=False,
    )
    pdf_path = pdf_dir / "SLA_Breach_Audit.pdf"
    if render.returncode != 0 or not pdf_path.is_file():
        return recalculated, "", (render.stderr or render.stdout or "LibreOffice PDF 渲染失败")[-1200:]
    extracted = subprocess.run(
        [pdftotext, "-raw", str(pdf_path), "-"],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )
    if extracted.returncode != 0:
        return recalculated, "", (extracted.stderr or "PDF 文本抽取失败")[-1200:]
    return recalculated, extracted.stdout, "LibreOffice 重算、另存和 PDF 渲染成功"


def read_rows(sheet, headers: list[str]) -> tuple[list[dict], list[str]]:
    values = list(sheet.iter_rows(values_only=True))
    if not values:
        return [], headers
    actual_headers = [str(value).strip() if value is not None else "" for value in values[0]]
    missing = [header for header in headers if header not in actual_headers]
    if missing:
        return [], missing
    positions = {header: actual_headers.index(header) for header in headers}
    records = []
    for row in values[1:]:
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        records.append({header: row[positions[header]] if positions[header] < len(row) else None for header in headers})
    return records, []


def formula_normalized(value) -> str:
    return re.sub(r"[\s$]", "", str(value or "")).upper()


def formulas_ok(workbook) -> tuple[bool, dict[str, str]]:
    analysis = workbook["Breach Analysis"]
    summary = workbook["Summary"]
    formulas = {f"Breach Analysis!F{row}": str(analysis[f"F{row}"].value or "") for row in range(2, 5)}
    formulas.update({f"Summary!B{row}": str(summary[f"B{row}"].value or "") for row in range(2, 5)})
    checks = []
    for row in range(2, 5):
        formula = formula_normalized(analysis[f"F{row}"].value)
        checks.append(formula.startswith("=") and f"E{row}" in formula and f"B{row}" in formula and "/" in formula and "100" in formula)
    total_formula = formula_normalized(summary["B2"].value)
    breached_formula = formula_normalized(summary["B3"].value)
    rate_formula = formula_normalized(summary["B4"].value)
    checks.extend(
        [
            total_formula.startswith("=") and "BREACHANALYSIS" in total_formula and "B2:B4" in total_formula,
            breached_formula.startswith("=") and "BREACHANALYSIS" in breached_formula and "E2:E4" in breached_formula,
            rate_formula.startswith("=") and "B3" in rate_formula and "B2" in rate_formula and "/" in rate_formula and "100" in rate_formula,
        ]
    )
    return all(checks), formulas


def evaluate(artifact_dir: Path) -> dict:
    workbook_path = artifact_dir / WORKBOOK_NAME
    email_path = artifact_dir / EMAIL_NAME
    results = []
    if not workbook_path.is_file() or not email_path.is_file():
        results.append(criterion("R001", False, f"Excel 存在={workbook_path.is_file()}，邮件草稿存在={email_path.is_file()}"))
        return build_result(results, [workbook_path, email_path])

    input_dir = Path(__file__).resolve().parents[1] / "data" / "input_files"
    expected_rows, expected_summary = derive_expected(input_dir)

    with tempfile.TemporaryDirectory(prefix="cowork-evo-sla-audit-") as temp_dir:
        recalculated_path, rendered_text, render_evidence = libreoffice_recalculate(workbook_path, Path(temp_dir))
        try:
            if not zipfile.is_zipfile(workbook_path):
                raise ValueError("交付文件不是有效 OOXML ZIP")
            formula_book = load_workbook(workbook_path, data_only=False)
            value_book = load_workbook(recalculated_path or workbook_path, data_only=True)
            sheet_names = value_book.sheetnames
            if "Breach Analysis" not in sheet_names or "Summary" not in sheet_names:
                raise ValueError(f"缺少指定工作表：{sheet_names}")
            actual_rows, analysis_missing = read_rows(value_book["Breach Analysis"], ANALYSIS_HEADERS)
            summary_rows, summary_missing = read_rows(value_book["Summary"], ["Metric", "Value"])
            formula_state, formula_evidence = formulas_ok(formula_book)
            analysis_format = formula_book["Breach Analysis"]
            summary_format = formula_book["Summary"]
            formatting_state = {
                "analysis_freeze": str(analysis_format.freeze_panes or ""),
                "summary_freeze": str(summary_format.freeze_panes or ""),
                "analysis_tables": len(analysis_format.tables),
                "summary_tables": len(summary_format.tables),
                "analysis_filter": str(analysis_format.auto_filter.ref or ""),
                "header_bold": all(cell.font.bold for cell in analysis_format[1]) and all(cell.font.bold for cell in summary_format[1]),
                "header_fill": all(cell.fill.fill_type == "solid" for cell in analysis_format[1]) and all(cell.fill.fill_type == "solid" for cell in summary_format[1]),
                "wide_columns": all((analysis_format.column_dimensions[column].width or 0) >= 14 for column in "ABCDEF"),
                "rate_format": all("0.00" in str(analysis_format[f"F{row}"].number_format) for row in range(2, 5)),
            }
            formula_book.close()
            value_book.close()
            email_text = email_path.read_text(encoding="utf-8")
        except Exception as error:
            results.append(criterion("R001", False, f"无法读取交付物：{error}"))
            return build_result(results, [workbook_path, email_path])

    structure_ok = (
        sheet_names == ["Breach Analysis", "Summary"]
        and not analysis_missing
        and not summary_missing
        and len(actual_rows) == len(expected_rows)
        and len(summary_rows) == len(SUMMARY_METRICS)
        and bool(email_text.strip())
    )
    results.append(
        criterion(
            "R001",
            structure_ok,
            f"工作表={sheet_names}，分析行={len(actual_rows)}/{len(expected_rows)}，摘要行={len(summary_rows)}/5，缺表头={analysis_missing + summary_missing}",
        )
    )

    priorities = [str(row["Priority"] or "").strip() for row in actual_rows]
    expected_priorities = [row["Priority"] for row in expected_rows]
    actual_targets = [row["SLA_Target_Hours"] for row in actual_rows]
    target_sort_ok = all(close(actual_targets[index], expected_rows[index]["SLA_Target_Hours"]) for index in range(min(len(actual_rows), len(expected_rows))))
    source_structure_ok = priorities == expected_priorities and target_sort_ok
    results.append(criterion("R002", source_structure_ok, f"优先级={priorities}，期望={expected_priorities}，SLA顺序={actual_targets}"))

    actual_by_priority = {str(row["Priority"] or "").strip(): row for row in actual_rows}
    metric_errors = []
    breach_errors = []
    rate_errors = []
    for expected in expected_rows:
        priority = expected["Priority"]
        actual = actual_by_priority.get(priority)
        if actual is None:
            metric_errors.append((priority, "缺失行"))
            breach_errors.append((priority, "缺失行"))
            rate_errors.append((priority, "缺失行"))
            continue
        if not (
            close(actual["Total_Tickets"], expected["Total_Tickets"], 0)
            and close(actual["Avg_Response_Hours"], expected["Avg_Response_Hours"], 0.001)
            and close(actual["SLA_Target_Hours"], expected["SLA_Target_Hours"], 0.001)
        ):
            metric_errors.append(
                (
                    priority,
                    actual["Total_Tickets"],
                    actual["Avg_Response_Hours"],
                    actual["SLA_Target_Hours"],
                    expected["Total_Tickets"],
                    float(expected["Avg_Response_Hours"]),
                    float(expected["SLA_Target_Hours"]),
                )
            )
        if not close(actual["Breached_Count"], expected["Breached_Count"], 0):
            breach_errors.append((priority, actual["Breached_Count"], expected["Breached_Count"]))
        if not close(actual["Breach_Rate"], expected["Breach_Rate"], 0.011):
            rate_errors.append((priority, actual["Breach_Rate"], float(expected["Breach_Rate"])))
    metric_ok = source_structure_ok and not metric_errors
    breach_ok = source_structure_ok and not breach_errors
    rate_ok = source_structure_ok and not rate_errors
    results.append(criterion("R003", metric_ok, f"错误={metric_errors}"))
    results.append(criterion("R004", breach_ok, f"错误={breach_errors}"))
    results.append(criterion("R005", rate_ok, f"错误={rate_errors}"))

    summary_names = [str(row["Metric"] or "").strip() for row in summary_rows]
    actual_summary = {str(row["Metric"] or "").strip(): row["Value"] for row in summary_rows}
    summary_errors = []
    for metric in SUMMARY_METRICS:
        if metric not in actual_summary:
            summary_errors.append((metric, "缺失", expected_summary[metric]))
        elif isinstance(expected_summary[metric], str):
            if normalized(actual_summary[metric]) != normalized(expected_summary[metric]):
                summary_errors.append((metric, actual_summary[metric], expected_summary[metric]))
        else:
            tolerance = 0.011 if metric == "Overall_Breach_Rate" else 0
            if not close(actual_summary[metric], expected_summary[metric], tolerance):
                summary_errors.append((metric, actual_summary[metric], float(expected_summary[metric])))
    summary_ok = summary_names == SUMMARY_METRICS and not summary_errors
    results.append(criterion("R006", summary_ok, f"指标顺序={summary_names}，错误={summary_errors}"))
    results.append(criterion("R007", formula_state, f"公式={formula_evidence}"))

    formatting_ok = (
        formatting_state["analysis_freeze"] == "A2"
        and formatting_state["summary_freeze"] == "A2"
        and formatting_state["analysis_tables"] >= 1
        and formatting_state["summary_tables"] >= 1
        and formatting_state["header_bold"]
        and formatting_state["header_fill"]
        and formatting_state["wide_columns"]
        and formatting_state["rate_format"]
    )
    results.append(criterion("R008", formatting_ok, json.dumps(formatting_state, ensure_ascii=False)))

    try:
        message = Parser(policy=policy.default).parsestr(email_text)
        body = str(message.get_payload() or "")
        rate_match = re.search(r"Overall_Breach_Rate\s*:\s*([0-9]+(?:\.[0-9]+)?)\s*%?", body, re.IGNORECASE)
        worst_match = re.search(r"Worst_Priority\s*:\s*([A-Za-z]+)", body, re.IGNORECASE)
        email_rate = float(rate_match.group(1)) if rate_match else None
        email_worst = worst_match.group(1) if worst_match else None
        email_ok = (
            normalized(message.get("From")) == "sla-audit@company.com"
            and normalized(message.get("To")) == "ops-director@company.com"
            and str(message.get("Subject") or "").strip() == "SLA Breach Audit Results"
            and close(email_rate, expected_summary["Overall_Breach_Rate"], 0.011)
            and normalized(email_worst) == normalized(expected_summary["Worst_Priority"])
            and summary_ok
        )
        email_evidence = f"From={message.get('From')}，To={message.get('To')}，Subject={message.get('Subject')}，违约率={email_rate}，最差={email_worst}"
    except Exception as error:
        email_ok = False
        email_evidence = f"邮件草稿解析失败：{error}"
    results.append(criterion("R009", email_ok, email_evidence))

    rendered_norm = normalized(rendered_text)
    rendered_compact = re.sub(r"[^a-z0-9_]+", "", rendered_norm)
    render_tokens = ["priority", "total_tickets", "overall_breach_rate", "high", "medium", "low", "31588", "24961", "7902"]
    missing_render_tokens = [token for token in render_tokens if token not in rendered_compact]
    render_ok = recalculated_path is not None and not missing_render_tokens and summary_ok and breach_ok and rate_ok
    results.append(criterion("R010", render_ok, f"{render_evidence}，渲染缺失={missing_render_tokens}"))

    core_error = not (source_structure_ok and metric_ok and breach_ok and rate_ok and summary_ok)
    broken_or_conflicting = not structure_ok or not email_ok or not render_ok
    results.append(criterion("R011", core_error, f"结构={source_structure_ok}，基础指标={metric_ok}，违约数={breach_ok}，违约率={rate_ok}，摘要={summary_ok}"))
    results.append(criterion("R012", broken_or_conflicting, f"文件结构={structure_ok}，邮件={email_ok}，LibreOffice={render_ok}"))
    return build_result(results, [workbook_path, email_path])


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
