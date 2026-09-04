#!/usr/bin/env python3
"""AutomationBench 5113 的 v2 分层负例：核心损坏归零，其余任务错误保留连续部分分。"""

from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
GOLDEN = ROOT / "solution" / "Golden"
INPUTS = ROOT / "environment" / "input_files"
OUTPUTS = (
    "I9_Compliance_Audit.xlsx",
    "I9_Compliance_Control_Brief.docx",
    "I9_Compliance_Review.pptx",
    "I9_Notification_Drafts.mbox",
    "I9_Tracker_Update_Drafts.json",
    "I9_Slack_Control.json",
    "I9_Execution_Control.json",
)
FATAL_CASES = {
    *(f"missing-{name}" for name in OUTPUTS),
    *(f"corrupt-{name}" for name in OUTPUTS),
    "mbox-empty",
    "unmodified-input-only",
}


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法载入 {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _unmodified_input_only(root: Path) -> None:
    for path in list(root.iterdir()):
        if path.is_file():
            path.unlink()
    for source in INPUTS.iterdir():
        if source.is_file():
            shutil.copy2(source, root / source.name)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    judge = _load("i9_5113_v2_judge", ROOT / "tests" / "judge.py")
    legacy = _load("i9_5113_v1_negatives", ROOT / "tests" / "negative_cases_v1.py")
    judge.V1.render = lambda path: ({".xlsx": 9, ".docx": 3, ".pptx": 3}[path.suffix.casefold()], "render-ok")
    cases = [(name, mutator) for name, mutator, _ in legacy.build_cases()]
    cases.append(("unmodified-input-only", _unmodified_input_only))
    if judge.evaluate(GOLDEN).get("reward") != 1.0:
        raise AssertionError("Golden v2 必须为1.0")

    results: list[dict] = []
    for name, mutator in cases:
        with tempfile.TemporaryDirectory(prefix=f"negative-5113-v2-{name}-") as temporary:
            output = Path(temporary) / "output"
            shutil.copytree(GOLDEN, output)
            mutator(output)
            result = judge.evaluate(output)
            reward = float(result.get("reward", 0.0))
            if name in FATAL_CASES:
                if reward != 0.0:
                    raise AssertionError(f"{name} 应硬归零，实际 {reward}")
                layer = "fatal_zero"
            else:
                if not (0.0 < reward < 1.0):
                    raise AssertionError(f"{name} 应获得0到1之间的连续部分分，实际 {reward}")
                layer = "partial"
            results.append({"case": name, "layer": layer, "reward": reward, "passed": bool(result.get("passed", False)), "penalties": result.get("penalties", [])})

    partial = [item["reward"] for item in results if item["layer"] == "partial"]
    report = {
        "total": len(results),
        "fatal_zero": sum(item["layer"] == "fatal_zero" for item in results),
        "partial_continuous": len(partial),
        "partial_reward_range": [min(partial), max(partial)],
        "cases": results,
    }
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
