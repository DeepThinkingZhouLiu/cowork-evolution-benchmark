#!/usr/bin/env python3
"""从两份原始 API JSON 独立重算真值的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import OrderedDict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-json-data-cleaning-pipeline"
RUBRIC_VERSION = "cowork-bench-json-data-cleaning-pipeline-v0.1"
JUDGE_VERSION = "cowork-bench-json-data-cleaning-pipeline-v0.1"
PASS_THRESHOLD = 0.8
OUTPUT_NAMES = [
    "clean_data.py",
    "cleaned_sales.json",
    "cleaned_inventory.json",
    "pipeline_results.json",
    "Data_Pipeline_Report.xlsx",
    "Pipeline_Dashboard.xlsx",
    "progress_note.json",
]
SALES_HEADERS = ["Product_Name", "Total_Quantity_Sold", "Total_Revenue", "Avg_Unit_Price"]
INVENTORY_HEADERS = ["Product_Name", "Current_Stock", "Reorder_Point", "Status"]
COMBINED_HEADERS = ["Product_Name", "Total_Sold", "Current_Stock", "Revenue", "Stock_Status"]
HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.3,
    "R003": 0.4,
    "R004": 0.45,
    "R005": 0.45,
    "R006": 0.45,
    "R007": 0.45,
    "R008": 0.5,
    "R009": 0.55,
    "R010": 0.6,
    "R011": 0.6,
    "R014": 0.7,
}
CRITERION_SPECS = {
    "R001": ("七个指定交付物均存在且可打开", 6, "hurdle"),
    "R002": ("清洗脚本在原始输入上可回放且三份 JSON 等于源数据真值", 9, "hurdle"),
    "R003": ("首条与重复条同时变异后脚本仍正确保留首条并联动重算", 10, "hurdle"),
    "R004": ("清洗销售恰有 18 条、无重复、规范商品名并保留其他字段", 7, "hurdle"),
    "R005": ("清洗库存恰有 5 条、无重复并保留俄文仓库字段", 7, "hurdle"),
    "R006": ("结果 JSON 的汇总、销售、库存和 Combined View 全部正确", 7, "hurdle"),
    "R007": ("报告 Excel 的三表、表头、行数与数据全部正确", 8, "hurdle"),
    "R008": ("库存 Status 为原生 IF 公式且 LibreOffice 重算正确", 6, "hurdle"),
    "R009": ("离线 Dashboard 与 Report 的 Combined View 完全一致", 6, "hurdle"),
    "R010": ("进度记录结构、动作与 18/5 清洗数量完整正确", 5, "hurdle"),
    "R011": ("三份 JSON、两份 Excel 与进度记录无核心事实冲突", 6, "hurdle"),
    "R012": ("两份 Excel 原生可编辑且达到专业交付要求", 4, "positive"),
    "R013": ("两份 Excel 可由 LibreOffice 重算、渲染并提取核心文本", 3, "positive"),
    "R014": ("进度记录没有无来源成本、业务、增长或质量效果断言", 5, "hurdle"),
    "R015": ("任一清洗记录、聚合或库存状态与源输入真值冲突", -16, "penalty"),
    "R016": ("清洗脚本不可回放、写死答案或未遵循保留首条语义", -16, "penalty"),
    "R017": ("交付缺失、损坏、公式写死、截图化或无法渲染", -16, "penalty"),
    "R018": ("交付物核心事实冲突或包含无来源业务效果断言", -18, "penalty"),
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


def load_records(path: Path) -> list[dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    records = payload.get("data") if isinstance(payload, dict) else payload
    if not isinstance(records, list) or not all(isinstance(row, dict) for row in records):
        raise ValueError(f"{path.name} data 不是对象数组")
    return records


def clean_first(records: list[dict[str, Any]], key: str) -> list[dict[str, Any]]:
    seen: set[str] = set()
    cleaned: list[dict[str, Any]] = []
    for source in records:
        identifier = str(source.get(key, ""))
        if not identifier or identifier in seen:
            continue
        seen.add(identifier)
        row = dict(source)
        row["product_name"] = str(row.get("product_name", "")).strip().title()
        cleaned.append(row)
    return cleaned


def derive_truth(input_dir: Path) -> dict[str, Any]:
    raw_sales = load_records(input_dir / "sales_api_response.json")
    raw_inventory = load_records(input_dir / "inventory_api_response.json")
    sales = clean_first(raw_sales, "order_id")
    inventory = clean_first(raw_inventory, "product_id")

    sales_agg: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for row in sales:
        name = str(row["product_name"])
        item = sales_agg.setdefault(name, {"quantity": 0, "revenue": 0.0, "prices": []})
        quantity = int(row["quantity"])
        price = float(row["unit_price"])
        item["quantity"] += quantity
        item["revenue"] += quantity * price
        item["prices"].append(price)
    sales_analysis = [
        {
            "Product_Name": name,
            "Total_Quantity_Sold": item["quantity"],
            "Total_Revenue": round(item["revenue"], 2),
            "Avg_Unit_Price": round(sum(item["prices"]) / len(item["prices"]), 2),
        }
        for name, item in sales_agg.items()
    ]

    inventory_status: list[dict[str, Any]] = []
    inventory_by_name: OrderedDict[str, dict[str, Any]] = OrderedDict()
    for row in inventory:
        status = "LOW" if int(row["current_stock"]) <= int(row["reorder_point"]) else "OK"
        item = {
            "Product_Name": row["product_name"],
            "Current_Stock": int(row["current_stock"]),
            "Reorder_Point": int(row["reorder_point"]),
            "Status": status,
        }
        inventory_status.append(item)
        inventory_by_name[str(row["product_name"])] = item

    names = list(sales_agg)
    names.extend(name for name in inventory_by_name if name not in sales_agg)
    combined = []
    for name in names:
        sales_item = sales_agg.get(name)
        inventory_item = inventory_by_name.get(name)
        combined.append({
            "Product_Name": name,
            "Total_Sold": sales_item["quantity"] if sales_item else 0,
            "Current_Stock": inventory_item["Current_Stock"] if inventory_item else 0,
            "Revenue": round(sales_item["revenue"], 2) if sales_item else 0.0,
            "Stock_Status": inventory_item["Status"] if inventory_item else "NO INVENTORY",
        })
    analysis = {
        "summary": {
            "raw_sales_count": len(raw_sales),
            "unique_sales_count": len(sales),
            "sales_duplicates_removed": len(raw_sales) - len(sales),
            "raw_inventory_count": len(raw_inventory),
            "unique_inventory_count": len(inventory),
            "inventory_duplicates_removed": len(raw_inventory) - len(inventory),
            "total_quantity_sold": sum(item["quantity"] for item in sales_agg.values()),
            "total_revenue": round(sum(item["revenue"] for item in sales_agg.values()), 2),
            "low_stock_products": sum(item["Status"] == "LOW" for item in inventory_status),
        },
        "sales_analysis": sales_analysis,
        "inventory_status": inventory_status,
        "combined_view": combined,
    }
    return {
        "cleaned_sales": {"sales": sales, "count": len(sales)},
        "cleaned_inventory": {"inventory": inventory, "count": len(inventory)},
        "analysis": analysis,
    }


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


def run_processor(script: Path, input_dir: Path) -> tuple[dict[str, Any] | None, str]:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-cleaner-") as directory:
        target = Path(directory)
        outputs = {
            "cleaned_sales": target / "cleaned_sales.json",
            "cleaned_inventory": target / "cleaned_inventory.json",
            "analysis": target / "pipeline_results.json",
        }
        command = [
            sys.executable,
            str(script.resolve()),
            "--sales", str((input_dir / "sales_api_response.json").resolve()),
            "--inventory", str((input_dir / "inventory_api_response.json").resolve()),
            "--sales-output", str(outputs["cleaned_sales"]),
            "--inventory-output", str(outputs["cleaned_inventory"]),
            "--analysis-output", str(outputs["analysis"]),
        ]
        try:
            completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=120, check=False)
        except Exception as error:
            return None, f"执行异常：{error}"
        if completed.returncode != 0 or not all(path.is_file() for path in outputs.values()):
            return None, f"退出={completed.returncode}，文件={[path.is_file() for path in outputs.values()]}，日志={completed.stdout[-1000:]}"
        try:
            return {
                key: json.loads(path.read_text(encoding="utf-8"))
                for key, path in outputs.items()
            }, completed.stdout[-1000:]
        except Exception as error:
            return None, f"JSON 无法读取：{error}"


def build_mutated_inputs(source: Path, target: Path) -> None:
    shutil.copytree(source, target)
    sales_path = target / "sales_api_response.json"
    sales_payload = json.loads(sales_path.read_text(encoding="utf-8"))
    for index, row in enumerate(sales_payload["data"]):
        if row["order_id"] == "ORD-001" and index == 0:
            row["quantity"] = 14
        if row["order_id"] == "ORD-003" and index > 17:
            row["quantity"] = 999
    sales_path.write_text(json.dumps(sales_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    inventory_path = target / "inventory_api_response.json"
    inventory_payload = json.loads(inventory_path.read_text(encoding="utf-8"))
    seen_p003 = False
    for row in inventory_payload["data"]:
        if row["product_id"] == "P004":
            row["current_stock"] = 30
        if row["product_id"] == "P003":
            if seen_p003:
                row["current_stock"] = 0
            seen_p003 = True
    inventory_path.write_text(json.dumps(inventory_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def table_rows(sheet: Any, headers: list[str], columns: int | None = None) -> list[dict[str, Any]] | None:
    count = columns or len(headers)
    actual_headers = [sheet.cell(1, column).value for column in range(1, count + 1)]
    if actual_headers != headers:
        return None
    rows: list[dict[str, Any]] = []
    for row_number in range(2, sheet.max_row + 1):
        values = [sheet.cell(row_number, column).value for column in range(1, count + 1)]
        if all(value in (None, "") for value in values):
            continue
        rows.append(dict(zip(headers, values)))
    return rows


def workbook_formatting(workbook: Any) -> dict[str, Any]:
    return {
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


def read_report(path: Path) -> tuple[dict[str, Any] | None, dict[str, Any], str]:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        return None, {}, str(error)
    formatting = workbook_formatting(workbook)
    if workbook.sheetnames != ["Sales Analysis", "Inventory Status", "Combined View"]:
        workbook.close()
        return None, formatting, f"工作表={workbook.sheetnames}"
    sales = table_rows(workbook["Sales Analysis"], SALES_HEADERS)
    inventory_base = table_rows(workbook["Inventory Status"], INVENTORY_HEADERS[:3], columns=3)
    combined = table_rows(workbook["Combined View"], COMBINED_HEADERS)
    formulas = [workbook["Inventory Status"][f"D{row}"].value for row in range(2, 7)]
    workbook.close()
    if sales is None or inventory_base is None or combined is None:
        return None, formatting, "表头不匹配"
    return {"sales": sales, "inventory_base": inventory_base, "combined": combined, "formulas": formulas}, formatting, ""


def read_dashboard(path: Path) -> tuple[list[dict[str, Any]] | None, dict[str, Any], str]:
    try:
        workbook = load_workbook(path, data_only=False)
    except Exception as error:
        return None, {}, str(error)
    formatting = workbook_formatting(workbook)
    if workbook.sheetnames != ["Pipeline Dashboard"]:
        workbook.close()
        return None, formatting, f"工作表={workbook.sheetnames}"
    rows = table_rows(workbook["Pipeline Dashboard"], COMBINED_HEADERS)
    workbook.close()
    return rows, formatting, "" if rows is not None else "表头不匹配"


def report_checks(payload: dict[str, Any] | None, truth: dict[str, Any]) -> tuple[bool, list[str]]:
    if not payload:
        return False, ["报告不可读取"]
    errors: list[str] = []
    analysis = truth["analysis"]
    if not payload_equal(payload["sales"], analysis["sales_analysis"]):
        errors.append("Sales Analysis 不匹配")
    expected_inventory = [
        {key: row[key] for key in INVENTORY_HEADERS[:3]}
        for row in analysis["inventory_status"]
    ]
    if not payload_equal(payload["inventory_base"], expected_inventory):
        errors.append("Inventory Status 基础列不匹配")
    if not payload_equal(payload["combined"], analysis["combined_view"]):
        errors.append("Combined View 不匹配")
    return not errors, errors


def recalculate_statuses(path: Path, target: Path) -> tuple[list[Any] | None, str]:
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
        statuses = [workbook["Inventory Status"][f"D{row}"].value for row in range(2, 7)]
        workbook.close()
        return statuses, completed.stdout[-500:]
    except Exception as error:
        return None, str(error)


def progress_checks(payload: Any, truth: dict[str, Any]) -> tuple[bool, list[str], str]:
    errors: list[str] = []
    if not isinstance(payload, dict) or set(payload) != {"entities", "relations"}:
        return False, ["顶层结构错误"], ""
    entities = payload.get("entities")
    if not isinstance(entities, list) or len(entities) != 1 or not isinstance(entities[0], dict):
        return False, ["entities 结构错误"], ""
    entity = entities[0]
    if entity.get("name") != "Data_Pipeline" or entity.get("entityType") != "project":
        errors.append("项目实体错误")
    observations = entity.get("observations")
    if not isinstance(observations, list) or not all(isinstance(item, str) for item in observations):
        return False, errors + ["observations 结构错误"], ""
    text = "\n".join(observations)
    lowered = text.lower()
    summary = truth["analysis"]["summary"]
    for token in (str(summary["unique_sales_count"]), str(summary["unique_inventory_count"])):
        if token not in text:
            errors.append(f"缺清洗数量 {token}")
    checks = [
        (("duplicate" in lowered or "去重" in text), "缺去重动作"),
        (("product_name" in lowered or "商品名" in text), "缺商品名规范化"),
        ("data_pipeline_report.xlsx" in lowered, "缺报告记录"),
        ("pipeline_dashboard.xlsx" in lowered, "缺离线 Dashboard 记录"),
    ]
    errors.extend(message for passed, message in checks if not passed)
    return not errors, errors, text


def unsupported_claims(text: str) -> list[str]:
    findings: list[str] = []
    lowered = text.lower()
    if re.search(r"\d+(?:\.\d+)?\s*%", text):
        findings.append("无来源百分比")
    patterns = [
        (r"cost savings?|saved?\s+(?:cost|money)|экономи[яил]", "成本节省"),
        (r"revenue (?:growth|increase|uplift)|increased revenue|рост выручки", "收入增长"),
        (r"improved? (?:quality|efficiency|performance)|quality improved|улучш", "质量或效率效果"),
        (r"business impact|roi|return on investment", "业务效果"),
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
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in artifact_paths if path.is_file()] or [hashlib.sha256(b"").hexdigest()]
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append({
            "criterion_id": item["criterion_id"],
            "score": score,
            "evidence": item["evidence"],
            "method": item["method"],
            "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]},
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

    truth = derive_truth(input_dir)
    replay, replay_log = run_processor(paths["clean_data.py"], input_dir)
    replay_ok = replay is not None and payload_equal(replay, truth)
    results.append(criterion("R002", replay_ok, f"匹配={replay_ok}，日志={replay_log}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-cleaning-mutation-") as directory:
        mutated_input = Path(directory) / "inputs"
        build_mutated_inputs(input_dir, mutated_input)
        mutated_truth = derive_truth(mutated_input)
        mutated_replay, mutation_log = run_processor(paths["clean_data.py"], mutated_input)
    mutation_changed = (
        mutated_truth["analysis"]["summary"]["total_quantity_sold"] == 188
        and close(mutated_truth["analysis"]["summary"]["total_revenue"], 3840)
        and mutated_truth["analysis"]["summary"]["low_stock_products"] == 1
        and next(row for row in mutated_truth["analysis"]["sales_analysis"] if row["Product_Name"] == "Gadget Pro")["Total_Quantity_Sold"] == 17
        and next(row for row in mutated_truth["analysis"]["inventory_status"] if row["Product_Name"] == "Gadget Pro")["Current_Stock"] == 45
    )
    mutation_ok = mutated_replay is not None and payload_equal(mutated_replay, mutated_truth) and mutation_changed
    results.append(criterion("R003", mutation_ok, f"匹配={mutated_replay is not None and payload_equal(mutated_replay, mutated_truth)}，真值变化={mutation_changed}，日志={mutation_log}"))

    delivered: dict[str, Any] = {}
    json_errors: list[str] = []
    for key, filename in (("cleaned_sales", "cleaned_sales.json"), ("cleaned_inventory", "cleaned_inventory.json"), ("analysis", "pipeline_results.json")):
        try:
            delivered[key] = json.loads(paths[filename].read_text(encoding="utf-8"))
        except Exception as error:
            delivered[key] = None
            json_errors.append(f"{filename}:{error}")
    sales_json_ok = payload_equal(delivered["cleaned_sales"], truth["cleaned_sales"])
    inventory_json_ok = payload_equal(delivered["cleaned_inventory"], truth["cleaned_inventory"])
    analysis_json_ok = payload_equal(delivered["analysis"], truth["analysis"])
    results.append(criterion("R004", sales_json_ok, f"匹配={sales_json_ok}，错误={json_errors}"))
    results.append(criterion("R005", inventory_json_ok, f"匹配={inventory_json_ok}，错误={json_errors}"))
    results.append(criterion("R006", analysis_json_ok, f"匹配={analysis_json_ok}，错误={json_errors}"))

    report_payload, report_formatting, report_error = read_report(paths["Data_Pipeline_Report.xlsx"])
    report_ok, report_errors = report_checks(report_payload, truth)
    results.append(criterion("R007", report_ok, f"错误={report_errors or report_error}"))
    formulas = report_payload.get("formulas", []) if report_payload else []
    formula_shape_ok = len(formulas) == 5 and all(
        isinstance(value, str)
        and value.startswith("=")
        and "IF" in value.upper()
        and re.search(rf"B{row}", value, re.IGNORECASE)
        and re.search(rf"C{row}", value, re.IGNORECASE)
        for row, value in enumerate(formulas, start=2)
    )
    with tempfile.TemporaryDirectory(prefix="cowork-evo-cleaning-recalc-") as directory:
        recalculated, recalc_log = recalculate_statuses(paths["Data_Pipeline_Report.xlsx"], Path(directory))
    expected_statuses = [row["Status"] for row in truth["analysis"]["inventory_status"]]
    recalc_ok = recalculated == expected_statuses
    formula_ok = formula_shape_ok and recalc_ok
    results.append(criterion("R008", formula_ok, f"公式={formulas}，重算={recalculated}，日志={recalc_log}"))

    dashboard_rows, dashboard_formatting, dashboard_error = read_dashboard(paths["Pipeline_Dashboard.xlsx"])
    dashboard_ok = dashboard_rows is not None and payload_equal(dashboard_rows, truth["analysis"]["combined_view"])
    results.append(criterion("R009", dashboard_ok, f"匹配={dashboard_ok}，错误={dashboard_error}"))

    try:
        progress_payload = json.loads(paths["progress_note.json"].read_text(encoding="utf-8"))
        progress_ok, progress_errors, progress_text = progress_checks(progress_payload, truth)
    except Exception as error:
        progress_payload = None
        progress_ok, progress_errors, progress_text = False, [str(error)], ""
    results.append(criterion("R010", progress_ok, f"错误={progress_errors}"))

    json_ok = sales_json_ok and inventory_json_ok and analysis_json_ok
    cross_ok = json_ok and report_ok and dashboard_ok and progress_ok
    results.append(criterion("R011", cross_ok, f"JSON={json_ok}，Report={report_ok}，Dashboard={dashboard_ok}，Progress={progress_ok}"))

    native_ok = (
        report_formatting.get("images") == 0
        and report_formatting.get("tables", 0) >= 3
        and report_formatting.get("freezes") == {"Sales Analysis": "A2", "Inventory Status": "A2", "Combined View": "A2"}
        and report_formatting.get("header_bold")
        and report_formatting.get("column_widths")
        and dashboard_formatting.get("images") == 0
        and dashboard_formatting.get("tables", 0) >= 1
        and dashboard_formatting.get("freezes") == {"Pipeline Dashboard": "A2"}
        and dashboard_formatting.get("header_bold")
        and dashboard_formatting.get("column_widths")
    )
    results.append(criterion("R012", native_ok, json.dumps({"report": report_formatting, "dashboard": dashboard_formatting}, ensure_ascii=False)))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-cleaning-render-") as directory:
        rendered, render_log = render_office([paths["Data_Pipeline_Report.xlsx"], paths["Pipeline_Dashboard.xlsx"]], Path(directory))
    report_render = rendered.get("Data_Pipeline_Report.xlsx", {})
    dashboard_render = rendered.get("Pipeline_Dashboard.xlsx", {})
    render_ok = (
        report_render.get("pages", 0) >= 3
        and all(token in report_render.get("text", "") for token in ("Product_Name", "Total_Quantity_Sold", "Current_Stock", "Stock_Status"))
        and dashboard_render.get("pages", 0) >= 1
        and all(token in dashboard_render.get("text", "") for token in ("Product_Name", "Total_Sold", "Revenue"))
    )
    results.append(criterion("R013", render_ok, f"页数=report:{report_render.get('pages', 0)},dashboard:{dashboard_render.get('pages', 0)}，日志={render_log}"))

    unsupported = unsupported_claims(progress_text)
    no_unsupported = not unsupported
    results.append(criterion("R014", no_unsupported, f"发现={unsupported}"))

    source_error = not (replay_ok and json_ok and report_ok and dashboard_ok and progress_ok)
    script_error = not (replay_ok and mutation_ok)
    damaged_error = not (structure_ok and report_payload is not None and dashboard_rows is not None and formula_ok and render_ok)
    conflict_or_unsupported = not cross_ok or bool(unsupported)
    results.append(criterion("R015", source_error, f"脚本={replay_ok}，JSON={json_ok}，Report={report_ok}，Dashboard={dashboard_ok}，Progress={progress_ok}"))
    results.append(criterion("R016", script_error, f"原始回放={replay_ok}，变异回放={mutation_ok}"))
    results.append(criterion("R017", damaged_error, f"结构={structure_ok}，Report={report_payload is not None}，Dashboard={dashboard_rows is not None}，公式={formula_ok}，渲染={render_ok}"))
    results.append(criterion("R018", conflict_or_unsupported, f"跨文件一致={cross_ok}，无来源断言={unsupported}"))
    return build_result(results, list(paths.values()))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir.resolve())
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
