#!/usr/bin/env python3
"""Independent deterministic Judge for the offline cash-flow workbook."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook

TASK_ID = "cowork-evo/automation-cashflow"
RUBRIC_VERSION = "automation-cashflow-v0.1"
OUTPUT = "cash_flow_forecast.xlsx"
INPUT = "cash_flow_inputs.xlsx"
EXPECTED_INPUT_SHA256 = "3c37565d46ea629b8b0963b6e52c1f906cd2bb41f2070b452b9b294cb7153527"
WEIGHTS = {"R001": .10, "R002": .22, "R003": .25, "R004": .15, "R005": .10, "R006": .13, "R007": .05, "R008": -.15}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def num(value: object) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    text = re.sub(r"[^0-9.\-]", "", str(value or ""))
    return float(text) if text else 0.0


def criterion(cid: str, ok: bool, evidence: str) -> dict:
    return {"criterion_id": cid, "score": 1.0 if ok else 0.0, "evidence": evidence, "method": "deterministic"}


def headers(ws):
    return [str(c.value).strip() if c.value is not None else "" for c in ws[1]]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--artifact-dir", type=Path, required=True)
    ap.add_argument("--result", type=Path, required=True)
    ap.add_argument("--reward-file", type=Path, required=True)
    args = ap.parse_args()
    root = Path(__file__).parents[1]
    input_path = root / "data/input_files" / INPUT
    output_path = args.artifact_dir / OUTPUT
    results: list[dict] = []
    try:
        native = output_path.is_file() and zipfile.is_zipfile(output_path)
        results.append(criterion("R001", native, "输出存在且为有效 OOXML"))
        if not native:
            results.extend([
                criterion("R002", False, "输出缺失或损坏"), criterion("R003", False, "输出缺失或损坏"),
                criterion("R004", False, "输出缺失或损坏"), criterion("R005", False, "输出缺失或损坏"),
                criterion("R006", False, "输出缺失或损坏"), criterion("R007", digest(input_path) == EXPECTED_INPUT_SHA256, "冻结输入 SHA256 检查"),
                criterion("R008", True, "无有效输出，惩罚项不触发"),
            ])
        else:
            wi = load_workbook(input_path, data_only=True)
            wo = load_workbook(output_path, data_only=True)
            required_source = ["Bank Balance", "AR Due", "AP Due"]
            sheet_ok = all(name in wo.sheetnames for name in required_source + ["Cash Flow Forecast", "Email Draft"])
            results[-1] = criterion("R001", sheet_ok, f"工作表={wo.sheetnames}")

            # Recompute truth from the visible input workbook.
            balance_h = {str(c.value).strip(): c.column for c in wi["Bank Balance"][1]}
            opening = num(wi["Bank Balance"].cell(2, balance_h["Balance"]).value)
            ar_h = {str(c.value).strip(): c.column for c in wi["AR Due"][1]}
            expected_rows: list[dict] = []
            for r in range(2, wi["AR Due"].max_row + 1):
                if str(wi["AR Due"].cell(r, ar_h["Status"]).value).strip() == "Include":
                    amount = num(wi["AR Due"].cell(r, ar_h["Amount"]).value); probability = num(wi["AR Due"].cell(r, ar_h["Collection Probability"]).value)
                    expected_rows.append({"type": "AR", "entity": str(wi["AR Due"].cell(r, ar_h["Customer"]).value).strip(), "source": amount, "probability": probability, "inflow": amount * probability / 100, "outflow": 0.0})
            ap_h = {str(c.value).strip(): c.column for c in wi["AP Due"][1]}
            for r in range(2, wi["AP Due"].max_row + 1):
                if str(wi["AP Due"].cell(r, ap_h["Status"]).value).strip() == "Include":
                    amount = num(wi["AP Due"].cell(r, ap_h["Amount"]).value)
                    expected_rows.append({"type": "AP", "entity": str(wi["AP Due"].cell(r, ap_h["Vendor"]).value).strip(), "source": amount, "probability": None, "inflow": 0.0, "outflow": amount})
            inflow = sum(row["inflow"] for row in expected_rows); outflow = sum(row["outflow"] for row in expected_rows); ending = opening + inflow - outflow

            forecast_ok = False; calculations_ok = False; excluded_ok = False
            if "Cash Flow Forecast" in wo:
                ws = wo["Cash Flow Forecast"]
                required = ["Type", "Entity", "Source Amount", "Collection Probability (%)", "Expected Inflow", "Outflow", "Status"]
                hm = {name: headers(ws).index(name) + 1 for name in required if name in headers(ws)}
                rows = []
                for r in range(2, ws.max_row + 1):
                    if ws.cell(r, 1).value is None:
                        break
                    rows.append({"type": str(ws.cell(r, hm["Type"]).value).strip(), "entity": str(ws.cell(r, hm["Entity"]).value).strip(), "source": num(ws.cell(r, hm["Source Amount"]).value), "probability": None if ws.cell(r, hm["Collection Probability (%)"]).value in (None, "") else num(ws.cell(r, hm["Collection Probability (%)"]).value), "inflow": num(ws.cell(r, hm["Expected Inflow"]).value), "outflow": num(ws.cell(r, hm["Outflow"]).value), "status": str(ws.cell(r, hm["Status"]).value).strip()})
                def same(a, b, key):
                    if key in {"type", "entity"}:
                        return a == b
                    if key == "probability":
                        return (a is None and b is None) or (a is not None and b is not None and abs(float(a) - float(b)) <= .01)
                    return abs(float(a) - float(b)) <= .01
                forecast_ok = bool(hm) and len(rows) == len(expected_rows) and all(all(same(o[key], e[key], key) for key in ("type", "entity", "source", "probability", "inflow", "outflow")) and o["status"] == "Include" for o, e in zip(rows, expected_rows))
                summary = {}
                for r in range(1, ws.max_row + 1):
                    label = ws.cell(r, 2).value
                    if label in {"Opening Balance", "Total Expected Inflows", "Total AP Outflows", "Ending Cash Balance"}:
                        summary[str(label)] = num(ws.cell(r, 3).value)
                calculations_ok = forecast_ok and all(abs(summary.get(k, float("nan")) - v) <= .01 for k, v in {"Opening Balance": opening, "Total Expected Inflows": inflow, "Total AP Outflows": outflow, "Ending Cash Balance": ending}.items())
                excluded_ok = not any(name in " ".join(str(c.value or "") for row in ws.iter_rows() for c in row) for name in ("Dispute Holdings", "Cancelled Lease"))
            results.append(criterion("R002", forecast_ok, f"预测行数={len(expected_rows)}"))
            results.append(criterion("R003", calculations_ok, f"opening={opening}; inflow={inflow}; outflow={outflow}; ending={ending}"))

            email_ok = False; body = ""
            if "Email Draft" in wo:
                ew = wo["Email Draft"]
                eh = {str(c.value).strip(): c.column for c in ew[1] if c.value is not None}
                vals = {str(ew.cell(r, eh["Field"]).value).strip(): ew.cell(r, eh["Value"]).value for r in range(2, ew.max_row + 1) if "Field" in eh and ew.cell(r, eh["Field"]).value is not None}
                body = str(vals.get("Body", ""))
                email_ok = vals.get("To") == "cfo@company.example.com" and all(row["entity"] in body for row in expected_rows) and all(f"{(row['inflow'] or row['outflow']):,.0f}" in body for row in expected_rows) and f"{ending:,.0f}" in body
            results.append(criterion("R004", email_ok, "Email Draft 收件人、实体和金额检查"))
            results.append(criterion("R005", excluded_ok and all(name not in body for name in ("Dispute Holdings", "Cancelled Lease", "173,400")), "排除行和错误摘要未进入交付"))

            source_preserved = sheet_ok
            if source_preserved:
                for name in required_source:
                    si, so = wi[name], wo[name]
                    if (si.max_row, si.max_column) != (so.max_row, so.max_column): source_preserved = False
                    for row in range(1, si.max_row + 1):
                        for col in range(1, si.max_column + 1):
                            if si.cell(row, col).value != so.cell(row, col).value: source_preserved = False
            style_ok = all(name in wo and wo[name].freeze_panes == "A2" and bool(wo[name].tables) and wo[name].auto_filter.ref for name in ("Cash Flow Forecast", "Email Draft")) if sheet_ok else False
            results.append(criterion("R006", source_preserved and style_ok, f"输入保留={source_preserved}; 结果格式={style_ok}"))
            hashes_ok = digest(input_path) == EXPECTED_INPUT_SHA256
            results.append(criterion("R007", hashes_ok, f"输入 SHA256={digest(input_path)}"))
            results.append(criterion("R008", False, "有效交付不触发空白/硬编码惩罚"))
            wi.close(); wo.close()
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
    by_id = {r["criterion_id"]: r["score"] for r in results}
    reward = sum(WEIGHTS[k] * by_id.get(k, 0.0) for k in WEIGHTS)
    passed = by_id.get("R001", 0) == 1 and all(by_id.get(k, 0) == 1 for k in ("R002", "R003", "R004", "R005")) and reward >= .8
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": "automation-cashflow-deterministic-v0.1", "artifact_sha256": [digest(output_path)] if output_path.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True); args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__": main()
