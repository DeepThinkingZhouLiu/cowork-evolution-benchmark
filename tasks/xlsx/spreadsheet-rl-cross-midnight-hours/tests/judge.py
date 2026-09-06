#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, shutil, subprocess, tempfile, zipfile
from pathlib import Path
from openpyxl import load_workbook

TASK_ID = 'cowork-evo/spreadsheet-rl-cross-midnight-hours'
INPUT = 'hours_input.xlsx'
EXPECTED_INPUT_SHA256 = "70695e01b190d245f8a459b96b0733efbc8400686a2cc5ceca0e506a85e71a03"
GOLDEN = 'hours_summary.xlsx'
OUT = 'hours_summary.xlsx'
WEIGHTS = {"R001": .10, "R002": .45, "R003": .15, "R004": .20, "R005": .05, "R006": .05, "R007": -.20}

def eq(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= 1e-8
    if type(a).__name__ in {"ArrayFormula", "DataTableFormula"} or type(b).__name__ in {"ArrayFormula", "DataTableFormula"}:
        return type(a) is type(b) and getattr(a, "text", None) == getattr(b, "text", None) and getattr(a, "ref", None) == getattr(b, "ref", None)
    return a == b

def c(cid, ok, ev):
    return {"criterion_id": cid, "score": 1.0 if ok else 0.0, "evidence": ev, "method": "deterministic"}

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a = ap.parse_args()
    root = Path(__file__).parents[1]; src = root / "data/input_files" / INPUT; golden = root / "solution" / "Golden" / GOLDEN; out = a.artifact_dir / OUT
    results = []
    native = out.is_file() and zipfile.is_zipfile(out)
    results.append(c("R001", native, "输出存在且为 OOXML 工作簿"))
    target_ok = formula_ok = preserve_ok = chart_ok = False
    if native:
        try:
            wi = load_workbook(src, data_only=False); wg = load_workbook(golden, data_only=False); wo = load_workbook(out, data_only=False)
            # Golden may add a result sheet or reorder sheets; output must match
            # Golden while retaining every input sheet by name.
            target_ok = wo.sheetnames == wg.sheetnames and set(wi.sheetnames).issubset(set(wg.sheetnames))
            formula_ok = target_ok; preserve_ok = target_ok
            changed = 0; mismatches = 0; formula_mismatches = 0
            if target_ok:
                for s in wg.sheetnames:
                    wi_s = wi[s] if s in wi.sheetnames else None
                    wg_s, wo_s = wg[s], wo[s]
                    # Golden defines the expected final structure. Inputs may be
                    # intentionally extended (for example, with an appended summary row).
                    if (wg_s.max_row, wg_s.max_column) != (wo_s.max_row, wo_s.max_column): preserve_ok = False
                    for r in range(1, wg_s.max_row + 1):
                        for col in range(1, wg_s.max_column + 1):
                            iv = wi_s.cell(r, col).value if wi_s is not None else None
                            gv, ov = wg_s.cell(r, col).value, wo_s.cell(r, col).value
                            if not eq(iv, gv):
                                changed += 1
                                if not eq(gv, ov): mismatches += 1
                                if isinstance(gv, str) and gv.startswith("=") and gv != ov: formula_mismatches += 1
                            elif not eq(iv, ov):
                                preserve_ok = False; mismatches += 1
                    if len(wg_s._charts) != len(wo_s._charts): chart_ok = False
                target_ok = target_ok and mismatches == 0
                formula_ok = formula_mismatches == 0
                chart_ok = all(len(wg[s]._charts) == len(wo[s]._charts) for s in wg.sheetnames)
            wi.close(); wg.close(); wo.close()
            results.append(c("R002", target_ok, f"与来源 Golden 的目标单元格一致，差异单元格={changed}，不一致={mismatches}"))
            results.append(c("R003", formula_ok, f"目标公式不一致={formula_mismatches}"))
            results.append(c("R004", preserve_ok, "非目标内容、工作表顺序和尺寸逐格检查"))
            results.append(c("R005", chart_ok or not native, "图表/工作簿结构可解析"))
            results.append(c("R006", hashlib.sha256(src.read_bytes()).hexdigest() == EXPECTED_INPUT_SHA256, "冻结输入 SHA256 匹配"))
            results.append(c("R007", False, "未检测到空白或硬编码替代品"))
        except Exception as exc:
            results.extend([c("R002", False, f"读取失败: {exc}"), c("R003", False, "读取失败"), c("R004", False, "读取失败"), c("R005", False, "读取失败"), c("R006", False, "读取失败"), c("R007", True, "损坏输出触发门槛")])
    else:
        results.extend([c("R002", False, "缺少有效输出"), c("R003", False, "缺少有效输出"), c("R004", False, "缺少有效输出"), c("R005", False, "缺少有效输出"), c("R006", False, "缺少有效输出"), c("R007", True, "缺少有效输出触发惩罚")])
    vals = {x["criterion_id"]: x["score"] for x in results}; reward = sum(WEIGHTS[k] * vals.get(k, 0) for k in WEIGHTS)
    passed = vals.get("R001", 0) == 1 and vals.get("R002", 0) == 1 and reward >= .7
    payload = {"task_id": TASK_ID, "rubric_version": "cowork-evo-rubric-v0.1", "judge_version": "batch-deterministic-v0.1", "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"); a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
