#!/usr/bin/env python3
"""从冻结 Rate Card/Actuals 输入独立重算 2026 预算并检查 XLSX 交付。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook


TASK_ID = "cowork-evo/automation-budget-prep"
RUBRIC_VERSION = "automation-budget-prep-static-v0.1"
JUDGE_VERSION = "automation-budget-prep-deterministic-v0.1"
OUTPUT_NAME = "Annual_Budget_2026.xlsx"
INPUT_NAME = "budget_preparation_source.xlsx"
HEADERS = ["Department", "2026 Salaries", "2026 Travel", "2026 Software"]
DEPARTMENTS = ["Engineering", "Sales", "Marketing"]
EXPECTED = {
    "Engineering": [1560000.0, 44000.0, 240000.0],
    "Sales": [832000.0, 165000.0, 60000.0],
    "Marketing": [520000.0, 66000.0, 120000.0],
}
TOTALS = [2912000.0, 275000.0, 420000.0]
WEIGHTS = {"R001": 0.10, "R002": 0.20, "R003": 0.25, "R004": 0.15, "R005": 0.15, "R006": 0.10, "R007": 0.05, "R008": -0.10}
EXPECTED_HASH = "e8437abfd82569d54761cbe45f348ae66013a61def0272cf780db450983f4f03"


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def criterion(cid: str, score: bool | float, evidence: str) -> dict:
    return {"criterion_id": cid, "score": float(score), "evidence": evidence, "method": "deterministic"}


def source_facts(path: Path) -> tuple[dict[str, list[float]], dict[str, float], dict[str, int]]:
    workbook = load_workbook(path, data_only=True)
    rates = {}
    for row in workbook["Rate Card"].iter_rows(min_row=2, values_only=True):
        if row[0]:
            rates[str(row[0])] = float(row[1])
    actuals = {}
    source_rows = {}
    for row_number, row in enumerate(workbook["2025 Actuals"].iter_rows(min_row=2, values_only=True), start=2):
        if row[0]:
            actuals[str(row[0])] = [float(row[1]), float(row[2]), float(row[3]), str(row[4])]
            source_rows[str(row[0])] = row_number
    workbook.close()
    expected = {}
    for department in DEPARTMENTS:
        values = actuals[department]
        expected[department] = [
            values[0] * (1 + rates["Salaries"] / 100),
            values[1] * (1 + rates["Travel"] / 100),
            values[2] * (1 + rates["Software"] / 100),
        ]
    return expected, rates, source_rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    root = Path(__file__).parents[1]
    input_path = root / "data/input_files" / INPUT_NAME
    output_path = args.artifact_dir / OUTPUT_NAME
    results: list[dict] = []
    try:
        expected, rates, source_rows = source_facts(input_path)
        valid_file = output_path.is_file() and zipfile.is_zipfile(output_path)
        if not valid_file:
            results.append(criterion("R001", False, "输出缺失或不是有效 OOXML 工作簿"))
            for cid in ("R002", "R003", "R004", "R005", "R006"):
                results.append(criterion(cid, False, "缺少有效输出"))
        else:
            workbook = load_workbook(output_path, data_only=False)
            expected_sheets = ["Rate Card", "2025 Actuals", "2026 Budget", "CFO Email Draft"]
            structure = workbook.sheetnames == expected_sheets
            results.append(criterion("R001", structure, f"工作表={workbook.sheetnames}"))
            budget = workbook["2026 Budget"]
            headers = [str(cell.value or "").strip() for cell in budget[1]]
            body_rows = []
            for row in budget.iter_rows(min_row=2, max_col=4, values_only=False):
                if row[0].value is None:
                    continue
                body_rows.append([row[0].value, row[1].value, row[2].value, row[3].value])
            departments = [str(row[0]).strip() for row in body_rows[:-1]] if body_rows else []
            rows_ok = headers == HEADERS and departments == DEPARTMENTS and "LegacyOps" not in departments and body_rows and str(body_rows[-1][0]).strip() == "Total"
            results.append(criterion("R002", rows_ok, f"表头={headers == HEADERS}; 部门={departments}; Total={bool(body_rows) and body_rows[-1][0] == 'Total'}"))
            values_ok = True
            formulas_ok = True
            for row in body_rows[:-1]:
                department = str(row[0]).strip()
                if department not in expected:
                    values_ok = False
                    formulas_ok = False
                    continue
                for index, target in enumerate(expected[department], start=1):
                    formula = row[index]
                    formulas_ok &= isinstance(formula, str) and formula.startswith("=") and "2025 Actuals" in formula and "Rate Card" in formula
                    # Validate formula shape against category-specific source row/rate reference.
                    expected_rate_row = index + 1
                    formulas_ok &= f"$B${expected_rate_row}" in str(formula)
                    formulas_ok &= "1+" in str(formula).replace(" ", "")
                    expected_source_col = "BCD"[index - 1]
                    formulas_ok &= f"!{expected_source_col}{source_rows.get(department, -1)}" in str(formula).replace("'", "")
                    if isinstance(formula, (int, float)):
                        values_ok &= abs(float(formula) - target) <= 0.01
            total_row = body_rows[-1] if body_rows else []
            totals_formula_ok = len(total_row) == 4 and all(isinstance(total_row[index], str) and total_row[index].startswith("=SUM(") for index in range(1, 4))
            totals_values_ok = all(abs(sum(expected[d][index] for d in DEPARTMENTS) - TOTALS[index]) <= 0.01 for index in range(3))
            results.append(criterion("R003", rows_ok and formulas_ok and totals_values_ok, f"类别增长率={rates}; 公式={formulas_ok}; 独立计算总计={totals_values_ok}"))
            results.append(criterion("R004", rows_ok and formulas_ok and totals_formula_ok, f"跨表公式={formulas_ok}; Total SUM={totals_formula_ok}"))
            email = workbook["CFO Email Draft"]
            fields = {str(row[0].value).strip(): row[1].value for row in email.iter_rows(min_row=2, max_col=2) if row[0].value}
            body = str(fields.get("Body", ""))
            email_ok = (
                fields.get("To") == "cfo@company.example.com"
                and "2026" in str(fields.get("Subject", ""))
                and fields.get("Status") == "Draft - Not Sent"
                and all(name in body for name in DEPARTMENTS)
                and all(token in body for token in ("2,912,000", "275,000", "420,000", "1,560,000", "44,000", "240,000", "832,000", "165,000", "60,000", "520,000", "66,000", "120,000"))
                and "LegacyOps" not in body
            )
            results.append(criterion("R005", email_ok, "CFO 草稿收件人、部门金额、合计和未发送状态检查"))
            professional = all(
                workbook[name].freeze_panes == "A2" and bool(workbook[name].tables) and workbook[name].auto_filter.ref
                for name in expected_sheets
            ) and all((budget.column_dimensions[c].width or 0) >= 12 for c in "ABCD")
            results.append(criterion("R006", professional, "四张表均有表格、筛选、冻结首行和可读列宽"))
            workbook.close()
        data_input = root / "data/input_files" / INPUT_NAME
        env_input = root / "environment/input_files" / INPUT_NAME
        trace_ok = digest(data_input) == EXPECTED_HASH and env_input.read_bytes() == data_input.read_bytes()
        results.append(criterion("R007", trace_ok, f"输入哈希={digest(data_input)}; data/environment 字节一致={env_input.read_bytes() == data_input.read_bytes()}"))
        violation = False
        if output_path.is_file() and zipfile.is_zipfile(output_path):
            check_wb = load_workbook(output_path, data_only=False)
            budget_values = [str(check_wb["2026 Budget"].cell(r, 1).value or "").strip() for r in range(2, check_wb["2026 Budget"].max_row + 1)] if "2026 Budget" in check_wb.sheetnames else []
            email_text = " ".join(str(v or "") for row in check_wb["CFO Email Draft"].iter_rows(values_only=True) for v in row) if "CFO Email Draft" in check_wb.sheetnames else ""
            violation = any(value == "LegacyOps" for value in budget_values) or "sent successfully" in email_text.lower()
            check_wb.close()
        results.append(criterion("R008", 1.0 if violation else 0.0, "完整性惩罚项；未检测到错误统一增长率、伪造发送或错误部门" if not violation else "检测到错误值或不应出现的内容"))
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
        for cid in ("R002", "R003", "R004", "R005", "R006", "R007", "R008"):
            results.append(criterion(cid, False, "读取失败"))
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[cid] * values.get(cid, 0.0) for cid in WEIGHTS)
    passed = reward >= 0.8 and all(values.get(cid, 0.0) == 1.0 for cid in ("R001", "R002", "R003", "R004", "R005"))
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest(output_path)] if output_path.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
