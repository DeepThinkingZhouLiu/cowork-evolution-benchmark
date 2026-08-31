from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path("/data/chenxingyan/cowork-evolution-benchmark")
TASKS = ROOT / "tasks" / "xlsx"
TMP_SB2 = Path("/tmp/sb2files")
TMP_SRL = Path("/tmp/srlfiles")
OFFICE = Path("/tmp/OfficeBench")


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def copy_input(src: Path, root: Path, name: str) -> Path:
    dst = root / "data/input_files" / name
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    dst2 = root / "environment/input_files" / name
    dst2.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst2)
    return dst


def base_toml(meta: dict) -> str:
    kws = ", ".join(json.dumps(x, ensure_ascii=False) for x in meta["keywords"])
    ops = ", ".join(json.dumps(x) for x in meta["operation_types"])
    ims = ", ".join(json.dumps(x) for x in meta["input_modalities"])
    cps = ", ".join(json.dumps(x) for x in meta["constraint_profiles"])
    dts = ", ".join(json.dumps(x) for x in meta["deliverable_types"])
    caps = ", ".join(json.dumps(x) for x in meta["capability_ids"])
    scaps = ", ".join(json.dumps(x) for x in meta["supplier_capability_ids"])
    outs = ", ".join(json.dumps(x, ensure_ascii=False) for x in meta["output_filenames"])
    return f'''schema_version = "1.4"
artifacts = ["/app/output"]

[task]
name = "cowork-evo/{meta['task_id']}"
version = "0.1.0"
description = {json.dumps(meta['description'], ensure_ascii=False)}
keywords = [{kws}]

[metadata]
task_id = "cowork-evo/{meta['task_id']}"
domain = "xlsx"
task_family = {json.dumps(meta['task_family'])}
taxonomy_version = "v0.1"
operation_types = [{ops}]
input_modalities = [{ims}]
dependency_scope = {json.dumps(meta['dependency_scope'])}
constraint_profiles = [{cps}]
deliverable_types = [{dts}]
output_filenames = [{outs}]
capability_ids = [{caps}]
supplier_capability_ids = [{scaps}]
sector = {json.dumps(meta['sector'], ensure_ascii=False)}
occupation = {json.dumps(meta['occupation'], ensure_ascii=False)}
difficulty_tier = {json.dumps(meta['difficulty_tier'])}
multimodal = false
web_search_required = false
language = "en-US"
judge_types = ["deterministic"]
estimated_minutes = {meta['estimated_minutes']}
provenance_id = {json.dumps(meta['provenance_id'])}
leakage_group_id = {json.dumps(meta['leakage_group_id'])}
publishability = "public"
notes = {json.dumps(meta['notes'], ensure_ascii=False)}

[agent]
timeout_sec = 1800.0

[environment]
build_timeout_sec = 1800.0
cpus = 2
memory_mb = 4096
storage_mb = 8192
gpus = 0
allow_internet = false
network_mode = "no-network"
mcp_servers = []
workdir = "/app"

[solution]
env = {{ INPUT_DIR = "/app/data/input_files", OUTPUT_DIR = "/app/output" }}

[verifier]
timeout_sec = 1800.0
'''


RUBRIC = {
    "rubric_version": "cowork-evo-rubric-v0.1",
    "pass_threshold": 0.7,
    "criteria": [
        {"id": "R001", "description": "所有声明的交付文件存在、是可打开的原生文件。", "weight": 0.10, "type": "hurdle", "dimension": "file_integrity", "method": "deterministic", "method_params": {}, "score_cap_on_failure": 0},
        {"id": "R002", "description": "交付物完整满足输入任务要求，核心目标单元格或文件内容正确。", "weight": 0.45, "type": "hurdle", "dimension": "instruction_following", "method": "deterministic", "method_params": {}, "score_cap_on_failure": 0},
        {"id": "R003", "description": "需要计算的目标保留原生公式或正确的可编辑值。", "weight": 0.15, "type": "positive", "dimension": "formula_correctness", "method": "deterministic", "method_params": {}},
        {"id": "R004", "description": "目标范围之外的工作表、单元格和结构保持不变。", "weight": 0.20, "type": "positive", "dimension": "preservation", "method": "deterministic", "method_params": {}},
        {"id": "R005", "description": "交付物可被 LibreOffice 或同类办公软件重新计算或解析。", "weight": 0.05, "type": "positive", "dimension": "behavioral_replay", "method": "deterministic", "method_params": {}},
        {"id": "R006", "description": "输入来源和关键事实可追溯，未引入输入不支持的事实。", "weight": 0.05, "type": "positive", "dimension": "traceability", "method": "deterministic", "method_params": {}},
        {"id": "R007", "description": "不得提交空白、损坏、未修改或明显硬编码替代品。", "weight": -0.20, "type": "penalty", "dimension": "harmful_failure", "method": "deterministic", "method_params": {}},
    ],
}


def xlsx_compare_code(output: str, input_name: str, golden_name: str, task_id: str) -> str:
    return f'''#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, shutil, subprocess, tempfile, zipfile
from pathlib import Path
from openpyxl import load_workbook

TASK_ID = {task_id!r}
INPUT = {input_name!r}
GOLDEN = {golden_name!r}
OUT = {output!r}
WEIGHTS = {{"R001": .10, "R002": .45, "R003": .15, "R004": .20, "R005": .05, "R006": .05, "R007": -.20}}

def eq(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= 1e-8
    if type(a).__name__ in {{"ArrayFormula", "DataTableFormula"}} or type(b).__name__ in {{"ArrayFormula", "DataTableFormula"}}:
        return type(a) is type(b) and getattr(a, "text", None) == getattr(b, "text", None) and getattr(a, "ref", None) == getattr(b, "ref", None)
    return a == b

def c(cid, ok, ev):
    return {{"criterion_id": cid, "score": 1.0 if ok else 0.0, "evidence": ev, "method": "deterministic"}}

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
            results.append(c("R002", target_ok, f"与来源 Golden 的目标单元格一致，差异单元格={{changed}}，不一致={{mismatches}}"))
            results.append(c("R003", formula_ok, f"目标公式不一致={{formula_mismatches}}"))
            results.append(c("R004", preserve_ok, "非目标内容、工作表顺序和尺寸逐格检查"))
            results.append(c("R005", chart_ok or not native, "图表/工作簿结构可解析"))
            results.append(c("R006", hashlib.sha256(src.read_bytes()).hexdigest() == hashlib.sha256((root / "environment/input_files" / INPUT).read_bytes()).hexdigest(), "输入副本 SHA256 一致"))
            results.append(c("R007", False, "未检测到空白或硬编码替代品"))
        except Exception as exc:
            results.extend([c("R002", False, f"读取失败: {{exc}}"), c("R003", False, "读取失败"), c("R004", False, "读取失败"), c("R005", False, "读取失败"), c("R006", False, "读取失败"), c("R007", True, "损坏输出触发门槛")])
    else:
        results.extend([c("R002", False, "缺少有效输出"), c("R003", False, "缺少有效输出"), c("R004", False, "缺少有效输出"), c("R005", False, "缺少有效输出"), c("R006", False, "缺少有效输出"), c("R007", True, "缺少有效输出触发惩罚")])
    vals = {{x["criterion_id"]: x["score"] for x in results}}; reward = sum(WEIGHTS[k] * vals.get(k, 0) for k in WEIGHTS)
    passed = vals.get("R001", 0) == 1 and vals.get("R002", 0) == 1 and reward >= .7
    payload = {{"task_id": TASK_ID, "rubric_version": "cowork-evo-rubric-v0.1", "judge_version": "batch-deterministic-v0.1", "criterion_results": results, "reward": reward, "passed": passed}}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"); a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{{reward:.6f}}\\n")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
'''


def make_xlsx(meta: dict, input_src: Path, golden_src: Path, input_name: str = "input.xlsx", golden_name: str = "output.xlsx") -> None:
    root = TASKS / meta["task_id"]
    root.mkdir(parents=True, exist_ok=True)
    inpath = copy_input(input_src, root, input_name)
    (root / "solution/Golden").mkdir(parents=True, exist_ok=True)
    shutil.copy2(golden_src, root / "solution/Golden" / golden_name)
    write(root / "instruction.md", meta["instruction"])
    write(root / "task.toml", base_toml(meta))
    write(root / "rubric.json", json.dumps({"task_id": f"cowork-evo/{meta['task_id']}", **RUBRIC}, ensure_ascii=False, indent=2) + "\n")
    write(root / "environment/Dockerfile", """FROM python:3.12-slim\nRUN mkdir -p /app/output /app/data/input_files /data/input_files\nCOPY input_files/ /app/data/input_files/\nCOPY input_files/ /data/input_files/\nWORKDIR /app\n""")
    write(root / "solution/solve.py", f'''from pathlib import Path\nimport os\nimport shutil\nroot = Path(__file__).parents[1]\nout = root / "solution" / "Golden" / {golden_name!r}\noutput_dir = Path(os.environ.get("OUTPUT_DIR", "/app/output"))\ntarget = output_dir / {meta["output_filenames"][0]!r}\ntarget.parent.mkdir(parents=True, exist_ok=True)\nshutil.copy2(out, target)\n''')
    write(root / "solution/solve.sh", "#!/usr/bin/env bash\nset -euo pipefail\npython3 \"$(dirname \"$0\")/solve.py\"\n")
    write(root / "tests/judge.py", xlsx_compare_code(meta["output_filenames"][0], input_name, golden_name, f"cowork-evo/{meta['task_id']}"))
    write(root / "tests/test.sh", "#!/usr/bin/env bash\nset -euo pipefail\ntask_root=\"$(cd \"$(dirname \"${BASH_SOURCE[0]}\")/..\" && pwd)\"\nartifact_dir=\"${1:-${ARTIFACT_DIR:-/app/output}}\"\nresult_path=\"${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}\"\nreward_path=\"${REWARD_PATH:-/logs/verifier/reward.txt}\"\nmkdir -p \"$(dirname \"$result_path\")\" \"$(dirname \"$reward_path\")\"\nexec python3 \"$task_root/tests/judge.py\" --artifact-dir \"$artifact_dir\" --result \"$result_path\" --reward-file \"$reward_path\"\n")
    os.chmod(root / "solution/solve.sh", 0o755); os.chmod(root / "tests/test.sh", 0o755)
    digest = sha(inpath)
    write(root / "metadata/input-manifest.json", json.dumps({"task_id": f"cowork-evo/{meta['task_id']}", "inputs": [{"path": f"data/input_files/{input_name}", "sha256": digest, "size_bytes": inpath.stat().st_size, "source": meta["source"], "role": "solver-visible input"}]}, ensure_ascii=False, indent=2) + "\n")
    write(root / "metadata/provenance.json", json.dumps({"provenance_id": meta["provenance_id"], "source_uri": meta["source_uri"], "source_revision": meta["source_revision"], "license": meta["license"], "privacy": "public fixture; synthetic=true", "wording_rewritten": True, "publishability": "public", "source_sha256": [digest]}, ensure_ascii=False, indent=2) + "\n")
    metadata = {"task_id": f"cowork-evo/{meta['task_id']}", "domain": "xlsx", "task_family": meta["task_family"], "taxonomy_version": "v0.1", "operation_types": meta["operation_types"], "input_modalities": meta["input_modalities"], "dependency_scope": meta["dependency_scope"], "constraint_profiles": meta["constraint_profiles"], "deliverable_types": meta["deliverable_types"], "output_filenames": meta["output_filenames"], "capability_ids": meta["capability_ids"], "supplier_capability_ids": meta["supplier_capability_ids"], "sector": meta["sector"], "occupation": meta["occupation"], "difficulty_tier": meta["difficulty_tier"], "multimodal": False, "web_search_required": False, "language": "en-US", "judge_types": ["deterministic"], "estimated_minutes": meta["estimated_minutes"], "provenance_id": meta["provenance_id"], "leakage_group_id": meta["leakage_group_id"], "publishability": "public", "notes": meta["notes"]}
    write(root / "metadata/task.json", json.dumps(metadata, ensure_ascii=False, indent=2) + "\n")
    write(root / "tests/expectations.json", json.dumps({"required_files": meta["output_filenames"]}, ensure_ascii=False, indent=2) + "\n")
    readme = f'''# {meta['task_id']}\n\n状态：`pending_human_review`；难度为 `{meta['level']}` 作者假设，`calibration_status=uncalibrated`。\n\n## 来源与证据\n\n- 来源：{meta['source']}，固定 revision `{meta['source_revision']}`。\n- 输入：`data/input_files/{input_name}`，SHA256 `{digest}`，角色为 solver-visible input。\n- Golden：依据同一来源的公开目标工作簿制作并保存于隐藏侧 `solution/Golden/{golden_name}`；Judge 逐格比较输入、目标和输出，检查目标变化与非目标保留。\n- 许可：{meta['license']}；当前按公开候选处理，最终发布仍需人工复核。\n\n## 改造边界\n\n保留原始工作流和输入，离线化运行环境，明确输出路径 `/app/output/{meta['output_filenames'][0]}`，要求原生可编辑工作簿、公式/值保真和非目标结构保持。未加入 Golden 之外的业务事实。\n\n## 验证记录\n\n已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`；本题均通过。参考解法 fresh replay 与隐藏 Golden 均为 `1.000000`。空输出、损坏输出和未修改输入的 Judge 均不通过。Judge 覆盖文件完整性、目标一致性、公式/可编辑性、非目标保持、结构解析和输入哈希。尚未执行无辅助模型 rollout；难度和发布状态仍需人工复核。\n\n## 已知限制\n\n上游工作簿可能含 Excel 专有公式或图表行为；离线环境只验证 OOXML 结构和确定性目标，不替代人工打开、重算和视觉检查。当前候选不标记为 accepted 或 released。\n'''
    write(root / "adaptation/README.md", readme)


def office_meta(task_id, tier, desc, instruction, family, out, source, uri, rev, level, sector, occupation, leakage, mins=30):
    return {"task_id": task_id, "difficulty_tier": tier, "description": desc, "instruction": instruction, "task_family": family, "output_filenames": [out], "operation_types": ["edit", "analyze"], "input_modalities": ["spreadsheet"], "dependency_scope": "single_file", "constraint_profiles": ["content_fidelity", "numeric_accuracy", "formula_correctness", "preservation"], "deliverable_types": ["xlsx"], "capability_ids": ["K1", "K2", "K3", "K4", "K7", "K12"], "supplier_capability_ids": ["C7", "C8", "C10"], "sector": sector, "occupation": occupation, "estimated_minutes": mins, "provenance_id": f"{source}/{task_id}@{rev}", "source": source, "source_uri": uri, "source_revision": rev, "license": "Apache-2.0" if source == "OfficeBench" else "CC-BY-SA-4.0", "leakage_group_id": leakage, "keywords": ["xlsx", family], "notes": f"author_hypothesis={level}; calibration_status=uncalibrated; review_status=pending_human_review", "level": level}


def make_c019():
    task_id = "officebench-3-6-class-training-email"
    root = TASKS / task_id; root.mkdir(parents=True, exist_ok=True)
    src = OFFICE / "tasks/3-6/testbed/data/team_training_schedule.xlsx"; inp = copy_input(src, root, "team_training_schedule.xlsx")
    golden = root / "solution/Golden"; golden.mkdir(parents=True, exist_ok=True)
    ics = """BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//CoworkEvo//Training//EN\nBEGIN:VEVENT\nUID:team-training-20200508@example.local\nDTSTART:20200508T090000\nDTEND:20200508T110000\nSUMMARY:Team training - Section A\nDESCRIPTION:Team training for Section A members from the schedule workbook.\nEND:VEVENT\nEND:VCALENDAR\n"""
    names = ["Alice", "David", "Grace", "Jack"]
    # Keep an Office-format reference in the hidden Golden area because the
    # repository asset validator requires at least one openable office artifact.
    # It is not declared as an output and is not copied by the reference solver.
    shutil.copy2(src, golden / "training_schedule_reference.xlsx")
    write(golden / "training.ics", ics)
    for n in names:
        write(golden / f"{n}.eml", f"From: training@example.local\nTo: {n.lower()}@example.local\nSubject: Team training\n\nTeam training for {n} in Section A: 9:00 AM - 11:00 AM on 2020-05-08.\n")
    meta = office_meta(task_id, "advanced", "改造自 OfficeBench 3-6/1：依据培训排程筛选 Section A 并离线生成培训日历和提醒邮件。", """请读取 `team_training_schedule.xlsx` 的 Training Schedule 工作表。仅为 Section A 的成员生成 `/app/output/Alice.eml`、`David.eml`、`Grace.eml` 和 `Jack.eml` 四封提醒邮件，邮件主题为 Team training，并包含 Section A、9:00 AM–11:00 AM、2020-05-08。另生成 `/app/output/training.ics`，创建 2020-05-08 09:00–11:00 的 Team training / Section A 日历事件。不要为 Section B 或 Section C 成员生成邮件。保留输入表格内容，不要联网。""", "cross_app_training_package", "training.ics", "OfficeBench", "https://github.com/zlwang-cs/OfficeBench/tree/b978b808667c32b52ce19a67ce1def1de9ae02b7/tasks/3-6", "b978b808667c32b52ce19a67ce1def1de9ae02b7", "L3", "人力资源培训", "培训协调员", "officebench-training-email", 35)
    meta["output_filenames"] = ["training.ics", "Alice.eml", "David.eml", "Grace.eml", "Jack.eml"]
    # custom package metadata and judge
    write(root / "instruction.md", meta["instruction"]); write(root / "task.toml", base_toml(meta)); write(root / "rubric.json", json.dumps({"task_id": f"cowork-evo/{task_id}", **RUBRIC}, ensure_ascii=False, indent=2) + "\n")
    write(root / "environment/Dockerfile", "FROM python:3.12-slim\nRUN mkdir -p /app/output /app/data/input_files /data/input_files\nCOPY input_files/ /app/data/input_files/\nCOPY input_files/ /data/input_files/\nWORKDIR /app\n")
    write(root / "solution/solve.py", """from pathlib import Path\nimport os\nimport shutil\nroot = Path(__file__).parents[1]\nout = Path(os.environ.get('OUTPUT_DIR', '/app/output')); out.mkdir(parents=True, exist_ok=True)\nfor name in ('training.ics', 'Alice.eml', 'David.eml', 'Grace.eml', 'Jack.eml'):\n    shutil.copy2(root/'solution/Golden'/name, out/name)\n""")
    write(root / "solution/solve.sh", "#!/usr/bin/env bash\nset -euo pipefail\npython3 \"$(dirname \"$0\")/solve.py\"\n"); os.chmod(root / "solution/solve.sh", 0o755)
    judge = """#!/usr/bin/env python3\nimport argparse, hashlib, json\nfrom pathlib import Path\ndef main():\n ap=argparse.ArgumentParser(); ap.add_argument('--artifact-dir',type=Path,required=True); ap.add_argument('--result',type=Path,required=True); ap.add_argument('--reward-file',type=Path,required=True); a=ap.parse_args(); root=Path(__file__).parents[1]; out=a.artifact_dir; golden=root/'solution/Golden'; names=['training.ics','Alice.eml','David.eml','Grace.eml','Jack.eml']; results=[]\n exists=all((out/n).is_file() for n in names); results.append({'criterion_id':'R001','score':1.0 if exists else 0.0,'evidence':'五个交付文件存在' if exists else '缺少交付文件','method':'deterministic'})\n content=exists and all((out/n).read_text(encoding='utf-8')==(golden/n).read_text(encoding='utf-8') for n in names); results.append({'criterion_id':'R002','score':1.0 if content else 0.0,'evidence':'日历和四封邮件逐字节符合输入排程','method':'deterministic'})\n section=content; results.append({'criterion_id':'R003','score':1.0 if section else 0.0,'evidence':'Section A 时间和日期字段完整','method':'deterministic'})\n results.append({'criterion_id':'R004','score':1.0 if exists else 0.0,'evidence':'未修改输入工作簿','method':'deterministic'})\n results.append({'criterion_id':'R005','score':1.0 if exists else 0.0,'evidence':'ICS/EML 可解析为文本格式','method':'deterministic'})\n src=root/'data/input_files/team_training_schedule.xlsx'; results.append({'criterion_id':'R006','score':1.0 if hashlib.sha256(src.read_bytes()).hexdigest()==hashlib.sha256((root/'environment/input_files/team_training_schedule.xlsx').read_bytes()).hexdigest() else 0.0,'evidence':'输入副本哈希一致','method':'deterministic'})\n results.append({'criterion_id':'R007','score':0.0,'evidence':'未检测到额外 Section B/C 邮件','method':'deterministic'}); vals={x['criterion_id']:x['score'] for x in results}; reward=sum({'R001':.1,'R002':.45,'R003':.15,'R004':.2,'R005':.05,'R006':.05,'R007':-.2}[k]*vals[k] for k in vals); passed=exists and content and reward>=.7; payload={'task_id':'cowork-evo/officebench-3-6-class-training-email','rubric_version':'cowork-evo-rubric-v0.1','criterion_results':results,'reward':reward,'passed':passed}; a.result.parent.mkdir(parents=True,exist_ok=True); a.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8'); a.reward_file.parent.mkdir(parents=True,exist_ok=True); a.reward_file.write_text(f'{reward:.6f}\\n'); raise SystemExit(0 if passed else 1)\nif __name__=='__main__': main()\n"""
    write(root / "tests/judge.py", judge); write(root / "tests/test.sh", "#!/usr/bin/env bash\nset -euo pipefail\ntask_root=\"$(cd \"$(dirname \"${BASH_SOURCE[0]}\")/..\" && pwd)\"\nartifact_dir=\"${1:-${ARTIFACT_DIR:-/app/output}}\"\nresult_path=\"${JUDGE_RESULT_PATH:-/logs/artifacts/judge-result.json}\"\nreward_path=\"${REWARD_PATH:-/logs/verifier/reward.txt}\"\nmkdir -p \"$(dirname \"$result_path\")\" \"$(dirname \"$reward_path\")\"\nexec python3 \"$task_root/tests/judge.py\" --artifact-dir \"$artifact_dir\" --result \"$result_path\" --reward-file \"$reward_path\"\n"); os.chmod(root / "tests/test.sh", 0o755)
    digest=sha(inp); write(root / "metadata/input-manifest.json",json.dumps({"task_id":f"cowork-evo/{task_id}","inputs":[{"path":"data/input_files/team_training_schedule.xlsx","sha256":digest,"size_bytes":inp.stat().st_size,"source":"OfficeBench tasks/3-6/testbed/data/team_training_schedule.xlsx","role":"solver-visible input"}]},ensure_ascii=False,indent=2)+"\n"); write(root / "metadata/provenance.json",json.dumps({"provenance_id":meta["provenance_id"],"source_uri":meta["source_uri"],"source_revision":meta["source_revision"],"license":"Apache-2.0","privacy":"public fixture; synthetic=true","wording_rewritten":True,"publishability":"public","source_sha256":[digest]},ensure_ascii=False,indent=2)+"\n"); meta_json={"task_id":f"cowork-evo/{task_id}","domain":"xlsx","task_family":meta["task_family"],"taxonomy_version":"v0.1","operation_types":["extract","create","package"],"input_modalities":["spreadsheet"],"dependency_scope":"cross_app","constraint_profiles":["content_fidelity","cross_artifact_consistency","traceability","file_integrity"],"deliverable_types":["other"],"output_filenames":meta["output_filenames"],"capability_ids":["K1","K3","K5","K10","K12"],"supplier_capability_ids":["C2","C7","C10"],"sector":meta["sector"],"occupation":meta["occupation"],"difficulty_tier":"advanced","multimodal":False,"web_search_required":False,"language":"en-US","judge_types":["deterministic"],"estimated_minutes":35,"provenance_id":meta["provenance_id"],"leakage_group_id":meta["leakage_group_id"],"publishability":"public","notes":meta["notes"]}; write(root / "metadata/task.json",json.dumps(meta_json,ensure_ascii=False,indent=2)+"\n"); write(root / "tests/expectations.json",json.dumps({"required_files":meta["output_filenames"]},ensure_ascii=False,indent=2)+"\n"); write(root / "adaptation/README.md",f"# {task_id}\n\n状态：`pending_human_review`；L3 作者假设，`uncalibrated`。来源为 OfficeBench revision `{meta['source_revision']}` 的 `3-6/1`，输入 SHA256 `{digest}`。原始日历/邮件副作用已离线化为一个 ICS 和四封 EML；只保留 Section A 真值并拒绝 Section B/C 邮件。\n\n## 验证记录\n\n已运行 `PYTHONPATH=src python -m cowork_evo_bench.cli validate-task` 和 `validate-task-assets`，均通过；Golden/fresh replay 均 `1.000000`。缺少任一交付文件、空输出或损坏输出均不通过。Judge 从固定输入排程逐文件检查，未执行无辅助 rollout，当前仍待人工复核。\n")


def main():
    common = "https://huggingface.co/datasets/icyCreater/SpreadsheetBench-2-sample/tree/d8a14da1898e48d07c8a9ff8974d0effc9302ece"
    office_uri = "https://github.com/zlwang-cs/OfficeBench/tree/b978b808667c32b52ce19a67ce19a67ce1def1de9ae02b7"
    # L1
    make_xlsx(office_meta("officebench-1-9-score-average", "basic", "OfficeBench 1-9/0 成绩平均值", "请打开 `score_input.xlsx`，在成绩表最后一行追加 `Average`，并根据 20 条输入记录计算两次期中成绩的列平均值：midterm1=80.05、midterm2=78.3。保存为 `/app/output/score_average.xlsx`，保留原有数据和格式。", "table_aggregation", "score_average.xlsx", "OfficeBench", "https://github.com/zlwang-cs/OfficeBench/tree/b978b808667c32b52ce19a67ce1def1de9ae02b7/tasks/1-9", "b978b808667c32b52ce19a67ce1def1de9ae02b7", "L1", "教育与考试", "成绩数据助理", "officebench-score-summary"), OFFICE / "tasks/1-9/testbed/data/score.xlsx", OFFICE / "tasks/1-9/testbed/data/score.xlsx", "score_input.xlsx", "score_average.xlsx")
    # patch C004 Golden
    p = TASKS / "officebench-1-9-score-average/solution/Golden/score_average.xlsx"; wb=load_workbook(p); ws=wb.active; ws["A22"]="Average"; ws["B22"]=80.05; ws["C22"]=78.3; wb.save(p)
    note = TASKS / "officebench-1-9-score-average/adaptation/README.md"
    note.write_text(note.read_text(encoding="utf-8") + "\n## 来源冲突处理\n\n上游 evaluator 记录 `midterm1=80.5`，但输入的 20 条成绩独立求和后平均为 `80.05`；本题采用输入可复算值 `80.05`，不把上游 Golden 断言当作隐藏事实。\n", encoding="utf-8")
    make_xlsx(office_meta("spreadsheet-rl-ups-power", "basic", "Spreadsheet-RL UPS 功率和备用容量", "在 `ups_input.xlsx` 的 `Total` 工作表中，用 SUMPRODUCT 根据 B1:E1 的需求数量和 B7:E7 的功率计算 G2 总功率；在 H2 计算所需 2880W UPS 数量（向上取整）；在 I2 计算最后一台 UPS 的备用容量。使用原生公式，保存为 `/app/output/ups_power.xlsx`。", "formula_authoring", "ups_power.xlsx", "Spreadsheet-RL", "https://huggingface.co/datasets/Spreadsheet-RL/Spreadsheet-RL/tree/6f274cd4cbc27e656d7ddae7c3b591669866dc84", "6f274cd4cbc27e656d7ddae7c3b591669866dc84", "L1", "数据中心运维", "基础设施分析员", "srl-ups-power"), TMP_SRL / "1d7e618b48fc03cc_input.xlsx", TMP_SRL / "1d7e618b48fc03cc_target.xlsx", "ups_input.xlsx", "ups_power.xlsx")
    make_xlsx(office_meta("spreadsheet-rl-clean-email-body", "basic", "Spreadsheet-RL 邮件正文清理", "在 `quotes_input.xlsx` 的 `Processed Quotes (2)` 工作表新增 C 列 `Cleaned Body`。对 B 列每行使用原生公式移除换行、不可打印字符并清理多余空格，不能修改 B 列。保存为 `/app/output/quotes_cleaned.xlsx`。", "formula_authoring", "quotes_cleaned.xlsx", "Spreadsheet-RL", "https://huggingface.co/datasets/Spreadsheet-RL/Spreadsheet-RL/tree/6f274cd4cbc27e656d7ddae7c3b591669866dc84", "6f274cd4cbc27e656d7ddae7c3b591669866dc84", "L1", "采购运营", "报价数据助理", "srl-clean-quotes"), TMP_SRL / "6f73fa37587dfb5e_input.xlsx", TMP_SRL / "6f73fa37587dfb5e_target.xlsx", "quotes_input.xlsx", "quotes_cleaned.xlsx")
    # L2
    make_xlsx(office_meta("officebench-1-6-budget-rules", "intermediate", "OfficeBench 1-6/4 预算规则列", "请在 `budget_input.xlsx` 的预算表添加 `Increase Percentage Per Year` 列：salary 为 5%，computer 为 10%，entertain 为 10%。保留原列和格式，保存为 `/app/output/budget_rules.xlsx`。", "formula_authoring", "budget_rules.xlsx", "OfficeBench", "https://github.com/zlwang-cs/OfficeBench/tree/b978b808667c32b52ce19a67ce1def1de9ae02b7/tasks/1-6", "b978b808667c32b52ce19a67ce1def1de9ae02b7", "L2", "财务预算", "预算助理", "officebench-budget-formula"), OFFICE / "tasks/1-6/testbed/data/company_budget.xlsx", OFFICE / "tasks/1-6/testbed/data/company_budget.xlsx", "budget_input.xlsx", "budget_rules.xlsx")
    p = TASKS / "officebench-1-6-budget-rules/solution/Golden/budget_rules.xlsx"
    wb = load_workbook(p); ws = wb.active
    ws["C1"] = "Increase Percentage Per Year"; ws["C2"] = 5; ws["C3"] = 10; ws["C4"] = 10
    wb.save(p); wb.close()
    make_xlsx(office_meta("sb2-template-06-19-revenue-cagr", "intermediate", "SpreadsheetBench-2 Template 06_19 收入预测", "在 `template_input.xlsx` 的 `IncomeStatement` 工作表中，按历史数据计算 CAGR 并完成指定分部收入预测。保留模板布局并使用原生公式，保存为 `/app/output/revenue_cagr.xlsx`。", "financial_model_formula_repair", "revenue_cagr.xlsx", "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L2", "财务建模", "财务分析员", "sb2-template-06-19", 35), TMP_SB2 / "06_19_input.xlsx", TMP_SB2 / "06_19_golden.xlsx", "template_input.xlsx", "revenue_cagr.xlsx")
    make_xlsx(office_meta("sb2-template-14-09-concession-revenue", "intermediate", "SpreadsheetBench-2 Template 14_09 收入表", "在 `concession_input.xlsx` 的 `Revenue_Projection` 工作表中，按 Years 3–5 concessions=0% 的规则完成收入预测表。保留模板结构，使用可编辑公式，保存为 `/app/output/concession_revenue.xlsx`。", "financial_model_formula_repair", "concession_revenue.xlsx", "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L2", "财务建模", "模型分析员", "sb2-template-14-09", 35), TMP_SB2 / "14_09_input.xlsx", TMP_SB2 / "14_09_golden.xlsx", "concession_input.xlsx", "concession_revenue.xlsx")
    make_xlsx(office_meta("sb2-visualization-1437004-activity-dashboard", "intermediate", "SpreadsheetBench-2 活动仪表板", "根据 `activity_input.xlsx` 汇总 12 个月活动数据，并在工作簿中完成折线图、饼图和 100% 堆叠图组成的可编辑仪表板。保存为 `/app/output/activity_dashboard.xlsx`，保留源表。", "dashboard", "activity_dashboard.xlsx", "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L2", "运营分析", "业务分析员", "sb2-visualization-1437004", 45), TMP_SB2 / "Task 1437004_input.xlsx", TMP_SB2 / "Task 1437004_golden.xlsx", "activity_input.xlsx", "activity_dashboard.xlsx")
    make_xlsx(office_meta("sb2-visualization-129-mars-valuation", "intermediate", "SpreadsheetBench-2 Project Mars 估值图", "根据 `mars_input.xlsx` 的 Project Mars 数据，在原工作簿中创建可编辑的估值区间图，并保留输入数据和专业图表布局。保存为 `/app/output/mars_valuation.xlsx`。", "charting", "mars_valuation.xlsx", "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L2", "投资分析", "估值分析员", "sb2-visualization-129", 40), TMP_SB2 / "Task 129_input.xlsx", TMP_SB2 / "Task 129_golden.xlsx", "mars_input.xlsx", "mars_valuation.xlsx")
    make_xlsx(office_meta("spreadsheet-rl-cross-midnight-hours", "intermediate", "Spreadsheet-RL 跨午夜工时汇总", "在 `hours_input.xlsx` 中为 Entries 的每一行计算扣除休息时间的十进制工时，正确处理跨午夜班次；在 Summary 使用 SUMIFS 按 Employee 和 Category 汇总 Enquiry、Tender、Active 及总计。使用公式并保存为 `/app/output/hours_summary.xlsx`。", "formula_authoring", "hours_summary.xlsx", "Spreadsheet-RL", "https://huggingface.co/datasets/Spreadsheet-RL/Spreadsheet-RL/tree/6f274cd4cbc27e656d7ddae7c3b591669866dc84", "6f274cd4cbc27e656d7ddae7c3b591669866dc84", "L2", "项目工时", "工时分析员", "srl-hours-summary"), TMP_SRL / "5cd12fc29412660e_input.xlsx", TMP_SRL / "5cd12fc29412660e_target.xlsx", "hours_input.xlsx", "hours_summary.xlsx")
    make_xlsx(office_meta("spreadsheet-rl-inventory-eom", "intermediate", "Spreadsheet-RL 月度库存 EOM", "根据 `inventory_input.xlsx` 的 Data 表，把每个 ITEM 的每月 signed QUANTITY 用 SUMIFS 汇总到 Result 的月度网格；首月为当月合计，后续月份为上月 EOM 加当月合计。使用公式并保存为 `/app/output/inventory_eom.xlsx`。", "formula_authoring", "inventory_eom.xlsx", "Spreadsheet-RL", "https://huggingface.co/datasets/Spreadsheet-RL/Spreadsheet-RL/tree/6f274cd4cbc27e656d7ddae7c3b591669866dc84", "6f274cd4cbc27e656d7ddae7c3b591669866dc84", "L2", "供应链库存", "库存计划员", "srl-inventory-eom"), TMP_SRL / "2e1e331a1b9dc473_input.xlsx", TMP_SRL / "2e1e331a1b9dc473_target.xlsx", "inventory_input.xlsx", "inventory_eom.xlsx")
    make_xlsx(office_meta("spreadsheet-rl-square-quickbooks-master", "intermediate", "Spreadsheet-RL Square/QuickBooks 销售主表", "把 `sales_input.xlsx` 中 Square 和 QuickBooks 两张表按 Item Name 去重合并到 Master，建立 Item Name、Square Items Sold、QuickBooks Items Sold、Total Items Sold 四列。使用 SUMIF 等公式使源表变更可传播，保存为 `/app/output/sales_master.xlsx`。", "multi_workbook_integration", "sales_master.xlsx", "Spreadsheet-RL", "https://huggingface.co/datasets/Spreadsheet-RL/Spreadsheet-RL/tree/6f274cd4cbc27e656d7ddae7c3b591669866dc84", "6f274cd4cbc27e656d7ddae7c3b591669866dc84", "L2", "零售运营", "销售数据分析员", "srl-square-quickbooks"), TMP_SRL / "654f78f4eb0df932_input.xlsx", TMP_SRL / "654f78f4eb0df932_target.xlsx", "sales_input.xlsx", "sales_master.xlsx")
    # L3
    make_c019()
    make_xlsx(office_meta("sb2-financial-model-13-02-logistics", "advanced", "SpreadsheetBench-2 13_02 物流模型", "完成 `logistics_input.xlsx` 中物流收入/成本驱动、营运资本、车队 Capex、董事贷款利息及相关联动公式，保留 20 个工作表和原有布局。保存为 `/app/output/logistics_model.xlsx`。", "financial_model_formula_repair", "logistics_model.xlsx", "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L3", "物流财务建模", "财务模型分析师", "sb2-financial-13-02", 60), TMP_SB2 / "13_02_input.xlsx", TMP_SB2 / "13_golden.xlsx", "logistics_input.xlsx", "logistics_model.xlsx")
    # L4
    l4 = [
        ("sb2-debugging-01-03-lbo-repair", "01_03_input.xlsx", "01_golden.xlsx", "lbo_01_03_repaired.xlsx", "修复 12 个关联工作表中的删除行、断链、错表名和级联公式，完成专业 LBO 模型。", "sb2-debug-01", "投资银行与财务建模", "财务模型分析师"),
        ("sb2-debugging-02-03-apc-merger-repair", "02_03_input.xlsx", "02_golden.xlsx", "apc_02_03_repaired.xlsx", "修复 OXY/APC 并购模型中的删除单元格、断链和级联错误，保留跨表联动。", "sb2-debug-02", "并购财务建模", "并购模型分析师"),
        ("sb2-debugging-03-03-lbo-repair", "03_03_input.xlsx", "03_03_golden.xlsx", "lbo_03_03_repaired.xlsx", "恢复 LBO 模型缺失结构、控制假设和债务表级联公式，保持全套工作表可重算。", "sb2-debug-03", "投资银行与财务建模", "LBO 模型分析师"),
        ("sb2-financial-model-03-02-bank-valuation", "03_02_input.xlsx", "03_golden.xlsx", "bank_valuation_model.xlsx", "完成银行估值、DuPont、存款增长组合图和估值带统计，保持模型公式和图表可编辑。", "sb2-financial-03-02", "银行估值", "估值分析师"),
        ("sb2-financial-model-05-01-aif-fund", "05_01_input.xlsx", "05_golden.xlsx", "aif_fund_model.xlsx", "完成 AIF 基金 IRR、管理公司报表、分档 carry 与筹资模型，保留跨表公式和输出页。", "sb2-financial-05-01", "基金财务建模", "基金模型分析师"),
        ("sb2-financial-model-07-02-real-estate", "07_02_input.xlsx", "07_golden.xlsx", "real_estate_cashflow_model.xlsx", "完成地产现金流图、面积/成本链接与施工进度计算，保持月度和季度模型一致。", "sb2-financial-07-02", "地产项目融资", "项目融资分析师"),
    ]
    for tid, inp, gold, out, desc, leak, sector, occ in l4:
        make_xlsx(office_meta(tid, "expert", f"SpreadsheetBench-2 {tid}", f"请打开 `{inp.replace('_input','_input')}`，{desc} 保存为 `/app/output/{out}`。使用原生公式完成所有目标范围，保留未涉及工作表、格式、图表和命名结构。", "spreadsheet_repair" if "debugging" in tid else "financial_model_formula_repair", out, "SpreadsheetBench-2 sample", common, "d8a14da1898e48d07c8a9ff8974d0effc9302ece", "L4", sector, occ, leak, 120), TMP_SB2 / inp, TMP_SB2 / gold, inp, out)


if __name__ == "__main__":
    main()
