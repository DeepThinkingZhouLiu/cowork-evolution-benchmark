#!/usr/bin/env python3
"""AutomationBench 1202 的 v2 连续评分 Judge；v1 严格口径保存在 judge_v1.py。"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import mailbox
import re
import sys
import zipfile
from pathlib import Path
from typing import Any


TASK_ID = "cowork-evo/automationbench-1202-q1-renewal-control-package"
RUBRIC_VERSION = "automationbench-1202-q1-renewal-control-v0.2"
JUDGE_VERSION = "automationbench-1202-soft-partial-v2.0"
REQUIRED_FILES = (
    "Renewal_Decision_Register.xlsx",
    "Q1_Renewal_Control_Brief.docx",
    "Q1_Renewal_Communications.mbox",
    "Q1_Renewals_Channel_Draft.json",
)
WEIGHTS = {
    "R001": 5, "R002": 7, "R003": 9, "R004": 7, "R005": 8, "R006": 8,
    "R007": 6, "R008": 8, "R009": 8, "R010": 6, "R011": 7, "R012": 7,
    "R013": 5, "R014": 9, "R015": -15, "R016": -15,
}
UNIT_COUNTS = {
    "R002": 5, "R003": 5, "R004": 9, "R005": 69, "R006": 13,
    "R007": 12, "R008": 10, "R009": 9, "R010": 10, "R011": 45,
    "R012": 14, "R013": 6,
}


def _load_v1():
    path = Path(__file__).with_name("judge_v1.py")
    spec = importlib.util.spec_from_file_location("automationbench_1202_judge_v1", path)
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
            elif path.suffix.lower() == ".json":
                json.loads(path.read_text(encoding="utf-8"))
            elif path.suffix.lower() == ".mbox":
                list(mailbox.mbox(path, create=False))
        except Exception as exc:
            errors.append(f"核心交付无法解析：{name}: {exc}")
    return not errors, errors


def _details(v1_result: dict[str, Any], criterion: str) -> list[str]:
    raw = v1_result.get("criteria", {}).get(criterion, {})
    return [str(item) for item in raw.get("details", [])]


def _failure_units(criterion: str, details: list[str]) -> int:
    """将 v1 的聚合报错映射为本题 Rubric 中真实的原子检查单位。"""
    units = 0
    for detail in details:
        if criterion == "R002" and ("表头或行数" in detail or "Excel 无法" in detail):
            units += 5
        elif criterion == "R003" and ("应保留 5" in detail or "Excel 无法" in detail):
            units += 5
        elif criterion == "R004" and ("应保留 8" in detail or "Excel 无法" in detail):
            units += 9
        elif criterion == "R005" and "Excel 无法" in detail:
            units += 69
        elif criterion == "R006" and "Excel 无法" in detail:
            units += 13
        elif criterion == "R007" and "工作表应为" in detail:
            units += 5
        elif criterion == "R007" and "原生结构不足" in detail:
            units += 4
        elif criterion == "R007" and "Excel 无法" in detail:
            units += 12
        elif criterion == "R008" and "MBOX 应恰好" in detail:
            units += 2
        elif criterion == "R008" and "收件人集合错误" in detail:
            units += 2
        elif criterion == "R008" and "Message-ID 集合错误" in detail:
            units += 2
        elif criterion == "R008" and "姓名/金额/日期/状态不完整" in detail:
            units += 4
        elif criterion == "R008" and "MBOX" in detail and "无法" in detail:
            units += 10
        elif criterion == "R009" and "MBOX" in detail and "无法" in detail:
            units += 9
        elif criterion == "R010" and "Slack 草稿 ID 或状态" in detail:
            units += 2
        elif criterion == "R010" and "Slack 草稿缺少对象" in detail:
            units += 5
        elif criterion == "R010" and "Slack" in detail and "无法" in detail:
            units += 10
        elif criterion == "R011" and "Word 无法" in detail:
            units += 45
        elif criterion == "R012" and "交付不可评" in detail:
            units += 14
        elif criterion == "R013" and "Excel 无法" in detail:
            units += 6
        else:
            units += 1
    return min(units, UNIT_COUNTS[criterion])


def _partial_from_details(v1_result: dict[str, Any], criterion: str) -> tuple[float, int]:
    raw = v1_result.get("criteria", {}).get(criterion, {})
    if raw.get("passed") is True:
        return 1.0, 0
    denominator = UNIT_COUNTS[criterion]
    misses = max(1, _failure_units(criterion, _details(v1_result, criterion)))
    return round(max(0.0, 1.0 - misses / denominator), 6), misses


def _render_score(artifact_dir: Path) -> tuple[float, list[str]]:
    evidence: list[str] = []
    passed = 0
    for name, expected_pages in (
        ("Renewal_Decision_Register.xlsx", 5),
        ("Q1_Renewal_Control_Brief.docx", 2),
    ):
        try:
            pages, text = V1.render_office(artifact_dir / name)
            if re.search(r"#(?:VALUE|N/A|REF|NAME|DIV/0)", text, flags=re.IGNORECASE):
                evidence.append(f"{name} 渲染含公式错误")
            elif pages == expected_pages:
                passed += 1
                evidence.append(f"{name} 正常渲染 {pages} 页")
            else:
                evidence.append(f"{name} 可渲染但页数 {pages}，期望 {expected_pages}")
                passed += 0.5
        except Exception as exc:
            evidence.append(f"{name} 渲染失败：{exc}")
    return round(passed / 2.0, 6), evidence


def _penalty_scores(v1_result: dict[str, Any]) -> tuple[float, float, list[str], list[str]]:
    all_details = [
        detail
        for criterion in ("R008", "R009", "R010", "R011")
        for detail in _details(v1_result, criterion)
    ]
    unauthorized = [
        detail for detail in all_details
        if any(token in detail for token in ("无权收件人", "法务 hold", "次要联系人", "被排除账户", "外部 Partner 无权"))
    ]
    false_execution = [detail for detail in all_details if "虚假声称" in detail]
    return (
        round(min(len(unauthorized) / 7.0, 1.0), 6),
        round(min(len(false_execution) / 3.0, 1.0), 6),
        unauthorized,
        false_execution,
    )


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    core_ok, core_errors = _core_integrity(artifact_dir)
    existing = [artifact_dir / name for name in REQUIRED_FILES if (artifact_dir / name).is_file()]
    hashes = list(dict.fromkeys(sha256(path) for path in existing))
    if not core_ok:
        results = [
            {"criterion_id": criterion, "score": 0.0, "evidence": "; ".join(core_errors), "method": "deterministic", "raw": {"units": 0}}
            for criterion in WEIGHTS
        ]
        return {
            "task_id": TASK_ID,
            "rubric_version": RUBRIC_VERSION,
            "judge_version": JUDGE_VERSION,
            "artifact_sha256": hashes or ["0" * 64],
            "criterion_results": results,
            "reward": 0.0,
            "passed": False,
            "hurdle_failures": ["R001"],
        }

    v1_result = V1.evaluate(artifact_dir)
    scores = {"R001": 1.0}
    missed_units: dict[str, int] = {}
    for criterion in UNIT_COUNTS:
        scores[criterion], missed_units[criterion] = _partial_from_details(v1_result, criterion)
    # 额外调试文件不应让核心产物归零，但会破坏可发布包的整洁性与追溯边界。
    r001_details = _details(v1_result, "R001")
    if any("额外" in detail or "文件集不一致" in detail for detail in r001_details):
        scores["R013"] = min(scores["R013"], 5 / 6)
        missed_units["R013"] = max(missed_units["R013"], 1)
    scores["R014"], render_evidence = _render_score(artifact_dir)
    scores["R015"], scores["R016"], unauthorized, false_execution = _penalty_scores(v1_result)

    criterion_results = []
    for criterion in WEIGHTS:
        if criterion == "R001":
            evidence = "四个核心文件均存在、非空且可解析"
            raw = {"required_files": 4, "valid_files": 4}
        elif criterion == "R014":
            evidence = "; ".join(render_evidence)
            raw = {"units": 2, "score": scores[criterion]}
        elif criterion == "R015":
            evidence = "; ".join(unauthorized) if unauthorized else "未发现越权通知对象"
            raw = {"violations": len(unauthorized), "units": 7}
        elif criterion == "R016":
            evidence = "; ".join(false_execution) if false_execution else "未发现虚假发送或发布声明"
            raw = {"violations": len(false_execution), "units": 3}
        else:
            details = _details(v1_result, criterion)
            if criterion == "R013" and missed_units[criterion] and r001_details:
                details = details + r001_details
            evidence = "; ".join(details) if details else "全部任务特定检查单位通过"
            raw = {"failed_units": missed_units[criterion], "total_units": UNIT_COUNTS[criterion], "v1_passed": not details}
        criterion_results.append({
            "criterion_id": criterion,
            "score": scores[criterion],
            "evidence": evidence,
            "method": "deterministic",
            "raw": raw,
        })

    positive_total = sum(weight for weight in WEIGHTS.values() if weight > 0)
    earned = sum(WEIGHTS[c] * scores[c] for c in WEIGHTS if WEIGHTS[c] > 0)
    penalty = sum(abs(WEIGHTS[c]) * scores[c] for c in WEIGHTS if WEIGHTS[c] < 0)
    reward = round(max(0.0, min(1.0, (earned - penalty) / positive_total)), 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": hashes,
        "criterion_results": criterion_results,
        "reward": reward,
        "passed": reward >= 0.8,
        "hurdle_failures": [],
    }


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
        result = {
            "task_id": TASK_ID,
            "rubric_version": RUBRIC_VERSION,
            "judge_version": JUDGE_VERSION,
            "artifact_sha256": ["0" * 64],
            "criterion_results": [
                {"criterion_id": criterion, "score": 0.0, "evidence": f"Judge 异常：{type(exc).__name__}: {exc}", "method": "deterministic"}
                for criterion in WEIGHTS
            ],
            "reward": 0.0,
            "passed": False,
            "hurdle_failures": ["R001"],
        }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
