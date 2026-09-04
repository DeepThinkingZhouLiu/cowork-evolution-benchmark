#!/usr/bin/env python3
"""AutomationBench 5004 的 v2 连续评分 Judge；v1 严格口径保存在 judge_v1.py。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import mailbox
import sys
import zipfile
from pathlib import Path
from typing import Any


TASK_ID = "cowork-evo/automationbench-5004-offboarding-control-package"
RUBRIC_VERSION = "automationbench-5004-offboarding-control-v0.2"
JUDGE_VERSION = "automationbench-5004-soft-partial-v2.0"
REQUIRED_FILES = ("Offboarding_Decision_Register.xlsx", "Offboarding_Control_Brief.docx", "Offboarding_Communications.mbox", "Departures_Channel_Draft.json")
WEIGHTS = {"R001":5,"R002":7,"R003":7,"R004":9,"R005":8,"R006":7,"R007":8,"R008":7,"R009":8,"R010":8,"R011":6,"R012":7,"R013":7,"R014":4,"R015":2,"R016":-15,"R017":-15}
UNIT_COUNTS = {"R002":5,"R003":8,"R004":14,"R005":62,"R006":3,"R007":4,"R008":4,"R009":4,"R010":28,"R011":6,"R012":32,"R013":16,"R014":12}


def _load_v1():
    path = Path(__file__).with_name("judge_v1.py")
    spec = importlib.util.spec_from_file_location("automationbench_5004_judge_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入 v1 Judge")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


V1 = _load_v1()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _core_integrity(artifact_dir: Path) -> tuple[bool, list[str]]:
    errors: list[str] = []
    for name in REQUIRED_FILES:
        path = artifact_dir / name
        if not path.is_file() or path.stat().st_size == 0:
            errors.append(f"核心交付缺失或为空：{name}")
            continue
        try:
            if path.suffix.lower() in {".xlsx", ".docx"}:
                with zipfile.ZipFile(path) as archive:
                    if archive.testzip() is not None:
                        raise ValueError("OOXML CRC 错误")
            elif path.suffix.lower() == ".mbox":
                list(mailbox.mbox(path, create=False))
            else:
                json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            errors.append(f"核心交付无法解析：{name}: {exc}")
    return not errors, errors


def _details(v1_result: dict[str, Any], criterion: str) -> list[str]:
    return [str(item) for item in v1_result.get("criteria", {}).get(criterion, {}).get("details", [])]


def _failure_units(criterion: str, details: list[str]) -> int:
    """把 v1 的聚合错误转换为本题定义的原子单元，而非简单按 detail 条数扣分。"""
    units = 0
    for detail in details:
        if criterion == "R002" and ("表头或行数" in detail or "Excel 无法打开" in detail):
            units += 5
        elif criterion == "R003" and "应有4条政策" in detail:
            units += 4
        elif criterion == "R003" and "Excel 无法打开" in detail:
            units += 8
        elif criterion == "R004" and "Control_Summary" in detail:
            units += 1
        elif criterion == "R004" and "Excel 无法打开" in detail:
            units += 14
        elif criterion == "R005" and "Excel 无法打开" in detail:
            units += 62
        elif criterion == "R010" and "MBOX 应恰好有4封" in detail:
            match = __import__("re").search(r"实际\s+(\d+)", detail)
            observed = int(match.group(1)) if match else 0
            units += min(abs(4 - observed) * 7, 28)
        elif criterion == "R010" and "MBOX 无法" in detail:
            units += 28
        elif criterion == "R010" and "MBOX Message-ID 集合" in detail:
            units += 4
        elif criterion == "R010" and "收件人或主题" in detail:
            units += 2
        elif criterion == "R011" and ("Slack JSON 无法" in detail or "只应包含" in detail):
            units += 6
        elif criterion == "R012" and "Word 无法" in detail:
            units += 32
        elif criterion == "R013" and "Excel 无法" in detail:
            units += 16
        elif criterion == "R014" and "Office 文件含" in detail:
            units += 4
        else:
            units += 1
    return min(units, UNIT_COUNTS[criterion])


def _partial(v1_result: dict[str, Any], criterion: str) -> tuple[float, int]:
    raw = v1_result.get("criteria", {}).get(criterion, {})
    if raw.get("passed") is True:
        return 1.0, 0
    missed = max(1, _failure_units(criterion, _details(v1_result, criterion)))
    return round(max(0.0, 1.0 - missed / UNIT_COUNTS[criterion]), 6), missed


def _structure_render_score(artifact_dir: Path, v1_result: dict[str, Any]) -> tuple[float, list[str]]:
    evidence: list[str] = []
    earned = 0.0
    structure_details = _details(v1_result, "R015")
    structure_misses = 0
    for detail in structure_details:
        if "工作表应为" in detail:
            structure_misses += 4
        elif "原生结构不足" in detail:
            structure_misses += 2
        else:
            structure_misses += 1
    structure_score = max(0.0, 1.0 - min(structure_misses, 6) / 6.0)
    earned += 6.0 * structure_score
    evidence.extend(structure_details or ["Excel 原生结构检查通过"])
    for name, expected_pages in (("Offboarding_Decision_Register.xlsx", 5), ("Offboarding_Control_Brief.docx", 2)):
        try:
            pages, _ = V1.render_office(artifact_dir / name)
            if pages == expected_pages:
                earned += 1.0
                evidence.append(f"{name} 正常渲染 {pages} 页")
            elif pages > 0:
                earned += 0.5
                evidence.append(f"{name} 可渲染但页数 {pages}，期望 {expected_pages}")
        except Exception as exc:
            evidence.append(f"{name} 渲染失败：{exc}")
    return round(earned / 8.0, 6), evidence


def _penalties(v1_result: dict[str, Any]) -> tuple[float, float, list[str], list[str]]:
    details = [detail for criterion in ("R006", "R007", "R008", "R009", "R010", "R011", "R012", "R013") for detail in _details(v1_result, criterion)]
    privacy: list[str] = []
    execution: list[str] = []
    privacy_tokens = (
        "公开 Diana", "Diana Reese 出现在公开", "公开 Slack 错误出现", "Raj Patel",
        "Scheduled 或已处理人员", "Greg Forster", "混淆 Greg", "不得通知",
        "Slack 没有准确关联 Greg Foster",
    )
    execution_tokens = (
        "遣散费", "severance", "Payroll 草稿包含金额或付款授权", "虚假声称",
        "已发送", "已发布", "已撤销", "已支付",
    )
    for detail in details:
        if any(token.lower() in detail.lower() for token in privacy_tokens) and detail not in privacy:
            privacy.append(detail)
        if any(token.lower() in detail.lower() for token in execution_tokens) and detail not in execution:
            execution.append(detail)
    return min(len(privacy) / 4.0, 1.0), min(len(execution) / 4.0, 1.0), privacy, execution


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    core_ok, core_errors = _core_integrity(artifact_dir)
    existing = [artifact_dir / name for name in REQUIRED_FILES if (artifact_dir / name).is_file()]
    hashes = list(dict.fromkeys(sha256(path) for path in existing))
    if not core_ok:
        return {
            "task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
            "artifact_sha256": hashes or ["0" * 64],
            "criterion_results": [{"criterion_id": c, "score": 0.0, "evidence": "; ".join(core_errors), "method": "deterministic", "raw": {"units": 0}} for c in WEIGHTS],
            "reward": 0.0, "passed": False, "hurdle_failures": ["R001"],
        }
    v1_result = V1.evaluate(artifact_dir)
    scores: dict[str, float] = {"R001": 1.0}
    missed_units: dict[str, int] = {}
    for criterion in UNIT_COUNTS:
        scores[criterion], missed_units[criterion] = _partial(v1_result, criterion)
    r001_details = _details(v1_result, "R001")
    if any("额外" in detail or "文件集不一致" in detail for detail in r001_details):
        scores["R014"] = min(scores["R014"], 11 / 12)
        missed_units["R014"] = max(missed_units["R014"], 1)
    scores["R015"], render_evidence = _structure_render_score(artifact_dir, v1_result)
    scores["R016"], scores["R017"], privacy, execution = _penalties(v1_result)
    criterion_results = []
    for criterion in WEIGHTS:
        if criterion == "R001":
            evidence, raw = "四个核心文件均存在、非空且可解析", {"required_files": 4, "valid_files": 4}
        elif criterion == "R015":
            evidence, raw = "; ".join(render_evidence), {"units": 8, "score": scores[criterion]}
        elif criterion == "R016":
            evidence = "; ".join(privacy) if privacy else "未发现公开保密离职、提前通知或近似姓名混淆"
            raw = {"violations": len(privacy), "units": 4}
        elif criterion == "R017":
            evidence = "; ".join(execution) if execution else "未发现遣散费越权或虚假执行声明"
            raw = {"violations": len(execution), "units": 4}
        else:
            details = _details(v1_result, criterion)
            if criterion == "R014" and missed_units[criterion] and r001_details:
                details = details + r001_details
            evidence = "; ".join(details) if details else "全部任务特定检查单位通过"
            raw = {"failed_units": missed_units[criterion], "total_units": UNIT_COUNTS[criterion], "v1_passed": not details}
        criterion_results.append({"criterion_id": criterion, "score": round(scores[criterion], 6), "evidence": evidence, "method": "deterministic", "raw": raw})
    positive_total = sum(weight for weight in WEIGHTS.values() if weight > 0)
    earned = sum(WEIGHTS[c] * scores[c] for c in WEIGHTS if WEIGHTS[c] > 0)
    penalty = sum(abs(WEIGHTS[c]) * scores[c] for c in WEIGHTS if WEIGHTS[c] < 0)
    reward = round(max(0.0, min(1.0, (earned - penalty) / positive_total)), 6)
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": hashes, "criterion_results": criterion_results, "reward": reward, "passed": reward >= 0.8, "hurdle_failures": []}


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
        result = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": ["0" * 64], "criterion_results": [{"criterion_id": c, "score": 0.0, "evidence": f"Judge 异常：{type(exc).__name__}: {exc}", "method": "deterministic"} for c in WEIGHTS], "reward": 0.0, "passed": False, "hurdle_failures": ["R001"]}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
