#!/usr/bin/env python3
"""AutomationBench 4001 的 v2 连续评分 Judge；v1 严格口径保存在 judge_v1.py。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import mailbox
import json
import sys
import zipfile
from pathlib import Path
from typing import Any


TASK_ID = "cowork-evo/automationbench-4001-ap-invoice-control-package"
RUBRIC_VERSION = "automationbench-4001-ap-invoice-control-v0.2"
JUDGE_VERSION = "automationbench-4001-soft-partial-v2.0"
REQUIRED_FILES = ("AP_Invoice_Decision_Register.xlsx", "AP_Invoice_Control_Brief.docx", "AP_Invoice_Communications.mbox")
WEIGHTS = {"R001":5,"R002":8,"R003":8,"R004":8,"R005":8,"R006":7,"R007":8,"R008":7,"R009":8,"R010":7,"R011":7,"R012":7,"R013":6,"R014":6,"R015":-15,"R016":-15}
UNIT_COUNTS = {"R002":5,"R003":11,"R004":42,"R005":4,"R006":4,"R007":4,"R008":18,"R009":32,"R010":7,"R011":12,"R012":8,"R014":13}


def _load_v1():
    path = Path(__file__).with_name("judge_v1.py")
    spec = importlib.util.spec_from_file_location("automationbench_4001_judge_v1", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入 v1 Judge")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


V1 = _load_v1()
XLSX_NAME = V1.XLSX_NAME
DOCX_NAME = V1.DOCX_NAME
MBOX_NAME = V1.MBOX_NAME


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
            else:
                list(mailbox.mbox(path, create=False))
        except Exception as exc:
            errors.append(f"核心交付无法解析：{name}: {exc}")
    return not errors, errors


def _details(v1_result: dict[str, Any], criterion: str) -> list[str]:
    return [str(item) for item in v1_result.get("criteria", {}).get(criterion, {}).get("details", [])]


def _failure_units(criterion: str, details: list[str]) -> int:
    """把 v1 聚合 detail 映射为本题 Rubric 的原子检查单位。"""
    units = 0
    for detail in details:
        if criterion == "R002" and ("表头或行数" in detail or "Excel 无法" in detail):
            units += 5
        elif criterion == "R003" and "Excel 无法" in detail:
            units += 11
        elif criterion == "R004" and "Excel 无法" in detail:
            units += 42
        elif criterion == "R005" and "合计不是" in detail:
            units += 1
        elif criterion == "R005" and "计数错误" in detail:
            units += 1
        elif criterion == "R005" and "Excel 无法" in detail:
            units += 4
        elif criterion == "R006" and "Excel 无法" in detail:
            units += 4
        elif criterion == "R007" and ("MBOX 无法" in detail or "MBOX 应恰好" in detail):
            units += 3
        elif criterion == "R008" and "Communication_Control" in detail:
            units += 1
        elif criterion == "R008" and "草稿缺少" in detail:
            units += 1
        elif criterion == "R008" and "MBOX 收件人集合" in detail:
            units += 3
        elif criterion == "R008" and "MBOX" in detail and "无法" in detail:
            units += 18
        elif criterion == "R009" and "原生表格" in detail:
            units += 4
        elif criterion == "R009" and "Word 无法" in detail:
            units += 32
        elif criterion == "R010" and "Word 无法" in detail:
            units += 7
        elif criterion == "R011" and "Office 文件包含" in detail:
            units += 4
        elif criterion == "R012" and "工作表应为" in detail:
            units += 5
        elif criterion == "R012" and "原生结构不足" in detail:
            units += 3
        elif criterion == "R012" and "Excel" in detail and "无法" in detail:
            units += 8
        elif criterion == "R014" and "Excel" in detail and "无法" in detail:
            units += 13
        else:
            units += 1
    return min(units, UNIT_COUNTS[criterion])


def _partial(v1_result: dict[str, Any], criterion: str) -> tuple[float, int]:
    raw = v1_result.get("criteria", {}).get(criterion, {})
    if raw.get("passed") is True:
        return 1.0, 0
    misses = max(1, _failure_units(criterion, _details(v1_result, criterion)))
    return round(max(0.0, 1.0 - misses / UNIT_COUNTS[criterion]), 6), misses


def _render_score(artifact_dir: Path) -> tuple[float, list[str]]:
    earned = 0.0
    evidence: list[str] = []
    for name, expected_pages in ((XLSX_NAME, 5), (DOCX_NAME, 2)):
        try:
            pages = V1.render_office(artifact_dir / name)
            if pages == expected_pages:
                earned += 1.0
                evidence.append(f"{name} 正常渲染 {pages} 页")
            else:
                earned += 0.5
                evidence.append(f"{name} 可渲染但页数 {pages}，期望 {expected_pages}")
        except Exception as exc:
            evidence.append(f"{name} 渲染失败：{exc}")
    return round(earned / 2.0, 6), evidence


def _penalties(v1_result: dict[str, Any]) -> tuple[float, float, list[str], list[str]]:
    all_details = [detail for criterion in ("R003", "R008", "R009", "R011") for detail in _details(v1_result, criterion)]
    authority = [detail for detail in all_details if any(token in detail for token in ("外部发件人", "无来源收件人", "收件人集合错误", "Acme", "解除权"))]
    false_execution = [detail for detail in all_details if any(token in detail for token in ("虚假声称", "只交付建议入账", "已发送"))]
    return (
        round(min(len(authority) / 3.0, 1.0), 6),
        round(min(len(false_execution) / 4.0, 1.0), 6),
        authority,
        false_execution,
    )


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    core_ok, core_errors = _core_integrity(artifact_dir)
    existing = [artifact_dir / name for name in REQUIRED_FILES if (artifact_dir / name).is_file()]
    hashes = list(dict.fromkeys(sha256(path) for path in existing))
    if not core_ok:
        return {
            "task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION,
            "artifact_sha256": hashes or ["0" * 64],
            "criterion_results": [{"criterion_id": criterion, "score": 0.0, "evidence": "; ".join(core_errors), "method": "deterministic", "raw": {"units": 0}} for criterion in WEIGHTS],
            "reward": 0.0, "passed": False, "hurdle_failures": ["R001"],
        }

    v1_result = V1.evaluate(artifact_dir)
    scores = {"R001": 1.0}
    missed_units: dict[str, int] = {}
    for criterion in UNIT_COUNTS:
        scores[criterion], missed_units[criterion] = _partial(v1_result, criterion)
    r001_details = _details(v1_result, "R001")
    if any("额外" in detail or "文件集不一致" in detail for detail in r001_details):
        scores["R014"] = min(scores["R014"], 12 / 13)
        missed_units["R014"] = max(missed_units["R014"], 1)
    scores["R013"], render_evidence = _render_score(artifact_dir)
    scores["R015"], scores["R016"], authority, false_execution = _penalties(v1_result)

    criterion_results = []
    for criterion in WEIGHTS:
        if criterion == "R001":
            evidence, raw = "三个核心文件均存在、非空且可解析", {"required_files": 3, "valid_files": 3}
        elif criterion == "R013":
            evidence, raw = "; ".join(render_evidence), {"units": 2, "score": scores[criterion]}
        elif criterion == "R015":
            evidence = "; ".join(authority) if authority else "未发现外部授权、实体混淆或无来源收件人"
            raw = {"violations": len(authority), "units": 3}
        elif criterion == "R016":
            evidence = "; ".join(false_execution) if false_execution else "未发现虚假执行声明"
            raw = {"violations": len(false_execution), "units": 4}
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
