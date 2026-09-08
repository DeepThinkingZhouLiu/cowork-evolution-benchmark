#!/usr/bin/env python3
"""用原三 Harness 配置重跑改造题，并按哈希复用未改题的 v0.2 重评分。"""

from __future__ import annotations

import argparse
import concurrent.futures
import json
import math
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import freeze_benchmark_release as release


LABELS = {"claude": "Claude Code 2.1.251", "opencode": "OpenCode 1.4.14", "openhands": "OpenHands SDK 1.44.1"}
OLD_DIRS = {"claude": "claude-code", "opencode": "opencode", "openhands": "openhands"}


def write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def build_plan(root, old_results):
    old_lock = read_json(root / "benchmarks/cowork-evo-240-v0.2/dataset.lock")
    new_lock = read_json(root / "benchmarks/cowork-evo-240-v0.3/dataset.lock")
    old = {x["instance_id"]: x for x in old_lock["tasks"]}
    current = [x for x in new_lock["tasks"] if x["instance_id"].startswith("docx/")]
    changed, reused = [], []
    for task in current:
        actual = release._task_content_sha256(root, root / task["task_path"])
        if actual != task["content_sha256"]:
            raise ValueError(f"任务不符合冻结快照：{task['instance_id']}")
        previous = old.get(task["instance_id"])
        (reused if previous and previous["content_sha256"] == actual else changed).append(task)
    if (len(changed), len(reused)) != (17, 43):
        raise ValueError(f"预期17题改造/43题复用，实际{len(changed)}/{len(reused)}")
    baseline, supplements = {}, {}
    for harness, directory in OLD_DIRS.items():
        data = read_json(old_results / directory / "results.json")
        rows = {row["task"]: row for row in data["tasks"]}
        baseline[harness] = []
        supplements[harness] = []
        for task in reused:
            slug = task["instance_id"].split("/", 1)[1]
            row = rows[slug]
            reward = row["new_reward"]
            if row["status"] != "completed" or not isinstance(reward, (int, float)) or not math.isfinite(reward) or not 0 <= reward <= 1:
                raise ValueError(f"旧结果不可复用：{harness}/{slug}")
            if any(not row.get(key) or not Path(row[key]).is_file()
                   for key in ("source_result_path", "judge_result_path")):
                supplements[harness].append(slug)
                continue
            baseline[harness].append({
                "task": slug, "reward": reward, "origin": "reused_v0.2_identical_task_hash",
                "content_sha256": task["content_sha256"],
                "source_result_path": row["source_result_path"],
                "judge_result_path": row["judge_result_path"],
            })
    return changed, baseline, supplements


def assert_snapshot(root, expected):
    actual = subprocess.check_output(["git", "rev-parse", "HEAD:tasks"], cwd=root, text=True).strip()
    dirty = subprocess.check_output(["git", "status", "--porcelain", "--", "tasks"], cwd=root, text=True)
    if actual != expected or dirty.strip():
        raise RuntimeError("任务快照发生变动，停止评测以避免混合版本")


def summarize(output, changed, baseline, frozen):
    panels = []
    for harness in LABELS:
        root = output / harness
        rows = list(baseline[harness])
        missing = []
        for index, task in changed[harness]:
            job = frozen.completed_job(root, index, task.name)
            if job is None:
                missing.append(task.name)
            else:
                record = frozen.attempt_record(job)
                rows.append({"task": task.name, "origin": "fresh_v0.3_rollout", **record})
        total = sum(row["reward"] for row in rows)
        item = {
            "harness": LABELS[harness], "valid_count": len(rows), "expected_count": 60,
            "fresh_valid_count": len(rows) - len(baseline[harness]), "reused_count": len(baseline[harness]),
            "reward_sum_valid": total, "mean_reward_valid": total / len(rows),
            "mean_reward_all_60": total / 60 if len(rows) == 60 else None,
            "nonzero_count": sum(row["reward"] > 0 for row in rows),
            "full_score_count": sum(row["reward"] >= 1 - 1e-9 for row in rows),
            "pending_or_infrastructure_failure_tasks": missing,
            "tasks": sorted(rows, key=lambda row: row["task"]),
        }
        write_json(root / "results.json", item)
        panels.append({k: v for k, v in item.items() if k != "tasks"})
    write_json(output / "panel-summary.json", {"updated_at": datetime.now(timezone.utc).isoformat(), "harnesses": panels})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--runner-root", type=Path, required=True)
    parser.add_argument("--old-results", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    parser.add_argument("--prepare-only", action="store_true")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    output = args.output_root.resolve()
    sys.path.insert(0, str(args.runner_root.resolve()))
    import run_office_module60 as office
    import run_cross_office60 as frozen

    office.configure_snapshot(root, "docx")
    frozen.SMOKE_TASK = "officeval-003-coordination-minutes-audit"
    changed_tasks, baseline, supplements = build_plan(root, args.old_results)
    changed_names = {x["instance_id"].split("/", 1)[1] for x in changed_tasks}
    changed = {harness: [(i, task) for i, task in frozen.ordered_tasks()
                        if task.name in changed_names or task.name in supplements[harness]]
               for harness in LABELS}
    tree = subprocess.check_output(["git", "rev-parse", "HEAD:tasks"], cwd=root, text=True).strip()
    assert_snapshot(root, tree)
    plan = {"task_tree": tree, "benchmark_root": str(root), "changed_tasks": changed_tasks,
            "new_rollouts": sum(map(len, changed.values())), "reused_results": sum(map(len, baseline.values())),
            "supplemental_reruns_missing_old_evidence": supplements, "provider": frozen.PROVIDER_BASE_URL,
            "model": frozen.CLAUDE_MODEL, "harnesses": LABELS,
            "protocol": "17题每Harness一次有效作答；未改题完整哈希相同且旧证据完整才复用v0.2重评分，否则补跑；基础设施失败最多3次尝试，不记0分；每Harness串行，总并发3。"}
    plan_path = output / "evaluation-plan.json"
    if plan_path.exists() and read_json(plan_path) != plan:
        raise ValueError("已有评测计划不匹配，必须使用新的输出目录")
    for harness in LABELS:
        frozen.write_manifest(output / harness, harness)
        office.amend_manifest(output / harness, "docx", warm_build_cache=False, build_proxy=None)
    write_json(plan_path, plan)
    summarize(output, changed, baseline, frozen)
    if args.prepare_only:
        print(f"预检通过：{plan['new_rollouts']}次新作答，{plan['reused_results']}份旧结果可复用。", flush=True)
        return 0
    environment = office.build_environment(frozen.load_provider_secret(frozen.SETTINGS))

    def run_pipeline(harness):
        write_json(output / harness / "status.json", {"state": "running"})
        for index, task in changed[harness]:
            assert_snapshot(root, tree)
            write_json(output / harness / "status.json", {"state": "running", "current_task": task.name})
            while frozen.latest_attempt(output / harness, index, task.name) < 3:
                status = frozen.run_one(harness, output / harness, index, task, environment, force=False)
                if status == 0:
                    break
            if frozen.completed_job(output / harness, index, task.name) is None:
                print(f"INFRA {harness} {task.name}: 已保留失败证据，继续后续题。", flush=True)
        write_json(output / harness / "status.json", {"state": "finished", "current_task": None})

    errors = {}
    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
        pending = {pool.submit(run_pipeline, name): name for name in LABELS}
        while pending:
            done, _ = concurrent.futures.wait(pending, timeout=15, return_when=concurrent.futures.FIRST_COMPLETED)
            for future in done:
                harness = pending.pop(future)
                try:
                    future.result()
                except Exception as exc:
                    errors[harness] = f"{type(exc).__name__}: {exc}"
                    write_json(output / harness / "status.json", {"state": "failed", "error": errors[harness]})
            summarize(output, changed, baseline, frozen)
    summary = read_json(output / "panel-summary.json")
    complete = all(x["valid_count"] == 60 for x in summary["harnesses"]) and not errors
    write_json(output / "status.json", {"state": "completed" if complete else "incomplete", "errors": errors})
    return 0 if complete else 2


if __name__ == "__main__":
    raise SystemExit(main())
