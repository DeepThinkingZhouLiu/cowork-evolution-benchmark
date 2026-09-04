#!/usr/bin/env python3
"""Independent deterministic Judge for the offline vendor reconciliation workbook."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook

TASK_ID = "cowork-evo/automation-vendor-recon"
RUBRIC_VERSION = "automation-vendor-recon-v0.1"
OUTPUT = "vendor_reconciliation.xlsx"
INPUT = "vendor_reconciliation_inputs.xlsx"
EXPECTED_INPUT_SHA256 = "d83a081ab8b53744c61726a252606f48eac758f0e13335a2846c3941fbeeb4a3"
WEIGHTS = {"R001": .10, "R002": .18, "R003": .27, "R004": .17, "R005": .10, "R006": .08, "R007": .05, "R008": -.15}


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


def hmap(ws):
    return {str(c.value).strip(): c.column for c in ws[1] if c.value is not None}


def main() -> None:
    ap = argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); args = ap.parse_args()
    root = Path(__file__).parents[1]; src = root / "data/input_files" / INPUT; env = root / "environment/input_files" / INPUT; out = args.artifact_dir / OUTPUT; results: list[dict] = []
    try:
        native = out.is_file() and zipfile.is_zipfile(out); results.append(criterion("R001", native, "输出存在且为有效 OOXML"))
        if not native:
            results.extend([criterion(k, False, "输出缺失或损坏") for k in ("R002", "R003", "R004", "R005", "R006")]); results.append(criterion("R007", digest(src) == EXPECTED_INPUT_SHA256 and digest(env) == EXPECTED_INPUT_SHA256, "输入哈希检查")); results.append(criterion("R008", True, "无有效输出，惩罚项不触发"))
        else:
            wi = load_workbook(src, data_only=True); wo = load_workbook(out, data_only=True)
            required_source = ["Vendor Statements Received", "Our AP Balances"]
            sheets_ok = all(n in wo.sheetnames for n in required_source + ["Reconciliation", "Email Draft"])
            results[-1] = criterion("R001", sheets_ok, f"工作表={wo.sheetnames}")
            st, apws = wi[required_source[0]], wi[required_source[1]]; sh, ah = hmap(st), hmap(apws)
            ap_data = {}
            for r in range(2, apws.max_row + 1): ap_data[str(apws.cell(r, ah["Vendor"]).value).strip()] = (num(apws.cell(r, ah["AP Balance"]).value), str(apws.cell(r, ah["Notes"]).value or "").strip())
            expected = []
            for r in range(2, st.max_row + 1):
                vendor = str(st.cell(r, sh["Vendor"]).value).strip(); email = str(st.cell(r, sh["Email"]).value).strip(); statement = num(st.cell(r, sh["Statement Balance"]).value); ap_balance, notes = ap_data[vendor]; excluded = "exclude from recon" in notes.lower(); diff = 0.0 if excluded else statement - ap_balance; status = "Excluded" if excluded else ("Matched" if abs(diff) < 1e-9 else "Unresolved discrepancy"); contact = "No" if excluded or abs(diff) < 1e-9 else "Yes"; expected.append({"vendor": vendor, "email": email, "statement": statement, "ap": ap_balance, "difference": diff, "status": status, "contact": contact, "notes": notes if excluded else ("Statement less AP balance" if contact == "Yes" else "")})
            recon_ok = False; calc_ok = False; exclusion_ok = False
            if "Reconciliation" in wo:
                rw = wo["Reconciliation"]; hm = hmap(rw); rows = []
                required = ["Vendor", "Email", "Statement Balance", "AP Balance", "Difference", "Status", "Contact Required", "Notes"]
                if all(k in hm for k in required):
                    for r in range(2, rw.max_row + 1):
                        rows.append({"vendor": str(rw.cell(r, hm["Vendor"]).value).strip(), "email": str(rw.cell(r, hm["Email"]).value).strip(), "statement": num(rw.cell(r, hm["Statement Balance"]).value), "ap": num(rw.cell(r, hm["AP Balance"]).value), "difference": num(rw.cell(r, hm["Difference"]).value), "status": str(rw.cell(r, hm["Status"]).value).strip(), "contact": str(rw.cell(r, hm["Contact Required"]).value).strip(), "notes": str(rw.cell(r, hm["Notes"]).value or "").strip()})
                    recon_ok = len(rows) == len(expected) and all(o["vendor"] == e["vendor"] and o["email"] == e["email"] and abs(o["statement"] - e["statement"]) <= .01 and abs(o["ap"] - e["ap"]) <= .01 for o, e in zip(rows, expected))
                    calc_ok = recon_ok and all(abs(o["difference"] - e["difference"]) <= .01 and o["status"] == e["status"] and o["contact"] == e["contact"] for o, e in zip(rows, expected))
                    exclusion_ok = calc_ok and rows[-1]["status"] == "Excluded" and rows[-1]["contact"] == "No" and rows[-1]["difference"] == 0
            results.append(criterion("R002", recon_ok, f"对账行数={len(expected)}")); results.append(criterion("R003", calc_ok, "差额、状态和联系规则独立重算"))
            body_all = ""; recipients = []
            if "Email Draft" in wo:
                ew = wo["Email Draft"]; eh = hmap(ew)
                for r in range(2, ew.max_row + 1):
                    recipients.append(str(ew.cell(r, eh["Recipient"]).value).strip()); body_all += "\n" + str(ew.cell(r, eh["Body"]).value or "")
            unresolved = [e for e in expected if e["contact"] == "Yes"]
            email_ok = "controller@company.example.com" in recipients and all(e["email"] in recipients and f"{abs(e['difference']):,.0f}" in body_all for e in unresolved) and all(e["vendor"] in body_all for e in unresolved)
            results.append(criterion("R004", email_ok, "汇总和未解决供应商邮件正文检查"))
            no_decoy = "billing@cloudhost.example.com" not in recipients and "billing@techserve.example.com" not in recipients and "CloudHost Pro" not in body_all and "TechServe Solutions" not in body_all
            results.append(criterion("R005", exclusion_ok and no_decoy, "已匹配/排除供应商未创建邮件"))
            preserved = sheets_ok
            if preserved:
                for name in required_source:
                    si, so = wi[name], wo[name]
                    if (si.max_row, si.max_column) != (so.max_row, so.max_column): preserved = False
                    for rr in range(1, si.max_row + 1):
                        for cc in range(1, si.max_column + 1):
                            if si.cell(rr, cc).value != so.cell(rr, cc).value: preserved = False
            style_ok = all(name in wo and wo[name].freeze_panes == "A2" and bool(wo[name].tables) and wo[name].auto_filter.ref for name in ("Reconciliation", "Email Draft")) if sheets_ok else False
            results.append(criterion("R006", preserved and style_ok, f"输入保留={preserved}; 结果格式={style_ok}"))
            hashes_ok = digest(src) == EXPECTED_INPUT_SHA256 and digest(env) == EXPECTED_INPUT_SHA256; results.append(criterion("R007", hashes_ok, f"输入 SHA256={digest(src)}")); results.append(criterion("R008", False, "有效交付不触发空白/硬编码惩罚")); wi.close(); wo.close()
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
    by_id = {r["criterion_id"]: r["score"] for r in results}; reward = sum(WEIGHTS[k] * by_id.get(k, 0.0) for k in WEIGHTS); passed = by_id.get("R001", 0) == 1 and all(by_id.get(k, 0) == 1 for k in ("R002", "R003", "R004", "R005")) and reward >= .8
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": "automation-vendor-recon-deterministic-v0.1", "artifact_sha256": [digest(out)] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}; args.result.parent.mkdir(parents=True, exist_ok=True); args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"); args.reward_file.parent.mkdir(parents=True, exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)


if __name__ == "__main__": main()
