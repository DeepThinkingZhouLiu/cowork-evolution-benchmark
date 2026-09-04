#!/usr/bin/env python3
"""Cowork-Bench MOEX 行业轮动题的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from openpyxl import load_workbook


OUTPUT_NAME = "sector_rotation_report.xlsx"
TASK_ID = "cowork-evo/cowork-bench-moex-sector-rotation-dashboard"
RUBRIC_VERSION = "cowork-bench-moex-sector-rotation-static-v0.1"
JUDGE_VERSION = "cowork-bench-moex-sector-rotation-static-v0.1"
HURDLE_CAPS = {"R001": 0.0, **{f"R{index:03d}": 0.74 for index in range(2, 8)}}
SYMBOLS = ["SBER.ME", "GAZP.ME", "LKOH.ME", "MGNT.ME", "MTSS.ME"]
PERFORMANCE_HEADERS = [
    "Symbol", "Sector", "Price_1Y_Ago", "Current_Price", "Return_Pct",
    "Benchmark_Return_Pct", "Alpha", "Analyst_Rating", "Target_Price", "Upside_Pct",
]
FINANCIAL_HEADERS = ["Symbol", "Revenue_Latest_Q", "Net_Income_Latest_Q", "Profit_Margin_Pct"]
SUMMARY_KEYS = [
    "Best_Performer", "Worst_Performer", "Avg_Alpha",
    "Stocks_Above_Benchmark", "Stocks_Below_Benchmark",
]

CRITERION_SPECS = {
    "R001": ("指定 XLSX 可打开且包含三张规定的可编辑工作表", 3, "hurdle"),
    "R002": ("五只股票顺序、行业、基准日和期末收盘价全部正确", 5, "hurdle"),
    "R003": ("五只股票的区间收益率全部正确", 4, "hurdle"),
    "R004": ("研究门户行业基准、评级和目标价全部正确联结", 5, "hurdle"),
    "R005": ("五只股票的 Alpha 和目标价上涨空间全部正确", 5, "hurdle"),
    "R006": ("五行年度营收、净利润和利润率全部正确", 6, "hurdle"),
    "R007": ("五项 Summary 指标全部正确", 5, "hurdle"),
    "R008": ("三张表的表头、固定顺序和可编辑性符合原 Query", 3, "positive"),
    "R009": ("工作簿可由 LibreOffice 重算并渲染出关键内容", 3, "positive"),
    "R010": ("任一来源价格、年报原值或门户字段错误", -10, "penalty"),
    "R011": ("任一核心派生指标或 Summary 错误", -10, "penalty"),
    "R012": ("交付损坏、截图化或无法渲染", -10, "penalty"),
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


def parse_portal(path: Path):
    source = path.read_text(encoding="utf-8")
    rows = [
        [html.unescape(re.sub(r"<[^>]+>", "", cell)).strip() for cell in re.findall(r"<td[^>]*>(.*?)</td>", row, flags=re.S | re.I)]
        for row in re.findall(r"<tr[^>]*>(.*?)</tr>", source, flags=re.S | re.I)
    ]
    rows = [row for row in rows if row]
    benchmarks = {row[0]: float(row[1]) for row in rows if len(row) == 2}
    ratings = {
        row[0]: {"sector": row[1], "rating": row[2], "target": float(row[3])}
        for row in rows if len(row) == 4
    }
    return benchmarks, ratings


def derive_expected(input_dir: Path):
    snapshot = json.loads((input_dir / "moex_snapshot.json").read_text(encoding="utf-8"))
    benchmarks, ratings = parse_portal(input_dir / "research_portal.html")
    prices = {(row["symbol"], row["date"]): float(row["close"]) for row in snapshot["prices"]}
    statements = {row["symbol"]: row for row in snapshot["annual_income_statements"]}
    performance = []
    financials = []
    for symbol in SYMBOLS:
        baseline = round(prices[(symbol, "2026-02-25")], 4)
        current = round(prices[(symbol, "2026-05-26")], 2)
        sector = ratings[symbol]["sector"]
        returned = round((current - baseline) / baseline * 100, 2)
        benchmark = benchmarks[sector]
        alpha = round(returned - benchmark, 2)
        target = ratings[symbol]["target"]
        upside = round((target - current) / current * 100, 2)
        performance.append(
            {
                "Symbol": symbol,
                "Sector": sector,
                "Price_1Y_Ago": baseline,
                "Current_Price": current,
                "Return_Pct": returned,
                "Benchmark_Return_Pct": benchmark,
                "Alpha": alpha,
                "Analyst_Rating": ratings[symbol]["rating"],
                "Target_Price": target,
                "Upside_Pct": upside,
            }
        )
        statement = statements[symbol]
        revenue = round(float(statement["total_revenue"]) / 1_000_000, 2)
        net_income = round(float(statement["net_income"]) / 1_000_000, 2)
        financials.append(
            {
                "Symbol": symbol,
                "Revenue_Latest_Q": revenue,
                "Net_Income_Latest_Q": net_income,
                "Profit_Margin_Pct": round(net_income / revenue * 100, 2),
            }
        )
    returns = {row["Symbol"]: row["Return_Pct"] for row in performance}
    alphas = {row["Symbol"]: row["Alpha"] for row in performance}
    summary = {
        "Best_Performer": max(returns, key=returns.get),
        "Worst_Performer": min(returns, key=returns.get),
        "Avg_Alpha": round(sum(alphas.values()) / len(alphas), 2),
        "Stocks_Above_Benchmark": sum(value > 0 for value in alphas.values()),
        "Stocks_Below_Benchmark": sum(value < 0 for value in alphas.values()),
    }
    return performance, financials, summary


def read_table(workbook, sheet_name: str, required_headers: list[str]):
    if sheet_name not in workbook.sheetnames:
        return [], required_headers, []
    rows = list(workbook[sheet_name].iter_rows(values_only=True))
    if not rows:
        return [], required_headers, []
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    indexes = {header: position for position, header in enumerate(headers)}
    missing = [header for header in required_headers if header not in indexes]
    if missing:
        return [], missing, headers
    records = []
    for row in rows[1:]:
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        records.append({header: row[indexes[header]] if indexes[header] < len(row) else None for header in required_headers})
    return records, [], headers


def read_summary(workbook):
    if "Summary" not in workbook.sheetnames:
        return {}, [], []
    rows = [row for row in workbook["Summary"].iter_rows(values_only=True) if any(value is not None for value in row)]
    keys = [str(row[0]).strip() if row else "" for row in rows]
    values = {key: row[1] if len(row) > 1 else None for key, row in zip(keys, rows)}
    return values, keys, rows


def libreoffice_prepare(path: Path, root: Path):
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
    source_copy = source_dir / OUTPUT_NAME
    shutil.copy2(path, source_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()
    process = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(source_copy)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    recalculated = recalc_dir / OUTPUT_NAME
    if process.returncode != 0 or not recalculated.is_file():
        return None, "", (process.stderr or process.stdout or "LibreOffice 重算失败")[-1000:]
    rendered = subprocess.run(
        [libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(recalculated)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    pdf_path = pdf_dir / f"{Path(OUTPUT_NAME).stem}.pdf"
    if rendered.returncode != 0 or not pdf_path.is_file():
        return recalculated, "", (rendered.stderr or rendered.stdout or "PDF 渲染失败")[-1000:]
    extracted = subprocess.run(
        [pdftotext, "-raw", str(pdf_path), "-"],
        capture_output=True, text=True, timeout=120, check=False,
    )
    if extracted.returncode != 0:
        return recalculated, "", (extracted.stderr or "PDF 文本提取失败")[-1000:]
    return recalculated, extracted.stdout, "LibreOffice 重算、重存和 PDF 渲染成功"


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []
    if not output_path.is_file():
        results.append(criterion("R001", False, f"缺少 {OUTPUT_NAME}"))
        return build_result(results, output_path)

    input_dir = Path(__file__).resolve().parents[1] / "data" / "input_files"
    expected_performance, expected_financials, expected_summary = derive_expected(input_dir)
    with tempfile.TemporaryDirectory(prefix="cowork-evo-moex-") as temp_dir:
        recalculated, rendered_text, render_evidence = libreoffice_prepare(output_path, Path(temp_dir))
        load_path = recalculated or output_path
        try:
            if not zipfile.is_zipfile(load_path):
                raise ValueError("不是有效 OOXML ZIP")
            workbook = load_workbook(load_path, data_only=True)
            sheet_names = workbook.sheetnames
            performance, performance_missing, performance_headers = read_table(workbook, "Performance", PERFORMANCE_HEADERS)
            financials, financial_missing, financial_headers = read_table(workbook, "Financials", FINANCIAL_HEADERS)
            summary, summary_keys, summary_rows = read_summary(workbook)
            workbook.close()
        except Exception as error:
            results.append(criterion("R001", False, f"XLSX 无法读取：{error}"))
            return build_result(results, output_path)

    structure_ok = (
        sheet_names == ["Performance", "Financials", "Summary"]
        and performance_headers == PERFORMANCE_HEADERS
        and financial_headers == FINANCIAL_HEADERS
        and summary_keys == SUMMARY_KEYS
        and len(performance) == 5
        and len(financials) == 5
        and len(summary_rows) == 5
        and not performance_missing
        and not financial_missing
    )
    results.append(
        criterion(
            "R001", structure_ok,
            f"工作表={sheet_names}，Performance 行={len(performance)}，Financials 行={len(financials)}，Summary键={summary_keys}，缺失表头={performance_missing + financial_missing}",
        )
    )

    source_errors = []
    return_errors = []
    portal_errors = []
    calculation_errors = []
    for index, expected in enumerate(expected_performance):
        if index >= len(performance):
            continue
        actual = performance[index]
        source_ok = (
            normalized(actual["Symbol"]) == normalized(expected["Symbol"])
            and normalized(actual["Sector"]) == normalized(expected["Sector"])
            and close(actual["Price_1Y_Ago"], expected["Price_1Y_Ago"], 0.00011)
            and close(actual["Current_Price"], expected["Current_Price"])
        )
        if not source_ok:
            source_errors.append((index + 2, actual, expected))
        if not close(actual["Return_Pct"], expected["Return_Pct"]):
            return_errors.append((expected["Symbol"], actual["Return_Pct"], expected["Return_Pct"]))
        portal_ok = (
            close(actual["Benchmark_Return_Pct"], expected["Benchmark_Return_Pct"])
            and normalized(actual["Analyst_Rating"]) == normalized(expected["Analyst_Rating"])
            and close(actual["Target_Price"], expected["Target_Price"])
        )
        if not portal_ok:
            portal_errors.append((expected["Symbol"], actual, expected))
        if not close(actual["Alpha"], expected["Alpha"]) or not close(actual["Upside_Pct"], expected["Upside_Pct"]):
            calculation_errors.append((expected["Symbol"], actual["Alpha"], actual["Upside_Pct"], expected["Alpha"], expected["Upside_Pct"]))

    exact_performance_rows = len(performance) == len(expected_performance)
    source_ok = exact_performance_rows and not source_errors
    returns_ok = exact_performance_rows and not return_errors
    portal_ok = exact_performance_rows and not portal_errors
    calculations_ok = exact_performance_rows and not calculation_errors
    results.append(criterion("R002", source_ok, f"来源错误={source_errors[:3]}"))
    results.append(criterion("R003", returns_ok, f"收益率错误={return_errors}"))
    results.append(criterion("R004", portal_ok, f"门户字段错误={portal_errors[:3]}"))
    results.append(criterion("R005", calculations_ok, f"Alpha/Upside错误={calculation_errors}"))

    financial_errors = []
    for index, expected in enumerate(expected_financials):
        if index >= len(financials):
            continue
        actual = financials[index]
        if not (
            normalized(actual["Symbol"]) == normalized(expected["Symbol"])
            and close(actual["Revenue_Latest_Q"], expected["Revenue_Latest_Q"])
            and close(actual["Net_Income_Latest_Q"], expected["Net_Income_Latest_Q"])
            and close(actual["Profit_Margin_Pct"], expected["Profit_Margin_Pct"])
        ):
            financial_errors.append((index + 2, actual, expected))
    financials_ok = len(financials) == len(expected_financials) and not financial_errors
    results.append(criterion("R006", financials_ok, f"财务数据错误={financial_errors[:3]}"))

    summary_errors = []
    for key, expected in expected_summary.items():
        actual = summary.get(key)
        if isinstance(expected, (int, float)) and not isinstance(expected, bool):
            if not close(actual, expected):
                summary_errors.append((key, actual, expected))
        elif normalized(actual) != normalized(expected):
            summary_errors.append((key, actual, expected))
    summary_ok = summary_keys == SUMMARY_KEYS and not summary_errors
    results.append(criterion("R007", summary_ok, f"Summary错误={summary_errors}"))

    results.append(criterion("R008", structure_ok, f"工作表、表头、行顺序和单元格结构通过={structure_ok}"))
    normalized_render = normalized(rendered_text)
    rendered_ok = (
        recalculated is not None
        and "sber.me" in normalized_render
        and "gazp.me" in normalized_render
        and ("127.3413" in normalized_render or "127,3413" in normalized_render)
    )
    results.append(criterion("R009", rendered_ok, f"{render_evidence}；渲染字符={len(normalized_render)}"))

    any_source_error = not (source_ok and portal_ok and financials_ok)
    any_calculation_error = not (returns_ok and calculations_ok and summary_ok)
    broken = not (structure_ok and rendered_ok)
    results.append(criterion("R010", any_source_error, f"价格/行业={source_ok}，门户={portal_ok}，年报={financials_ok}"))
    results.append(criterion("R011", any_calculation_error, f"收益率={returns_ok}，Alpha/Upside={calculations_ok}，Summary={summary_ok}"))
    results.append(criterion("R012", broken, f"结构={structure_ok}，渲染={rendered_ok}，{render_evidence}"))
    return build_result(results, output_path)


def build_result(results: list[dict], artifact_path: Path) -> dict:
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
    digest = hashlib.sha256(artifact_path.read_bytes()).hexdigest() if artifact_path.is_file() else hashlib.sha256(b"").hexdigest()

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
        "artifact_sha256": [digest],
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= 0.75,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
