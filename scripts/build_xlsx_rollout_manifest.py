#!/usr/bin/env python3
"""将 Excel rollout 汇总整理为可提交的题级 JSON 和 CSV 清单。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import statistics
from pathlib import Path


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_metadata(task_dir: Path) -> dict:
    metadata_path = task_dir / "metadata" / "task.json"
    return json.loads(metadata_path.read_text(encoding="utf-8"))


def build_rows(summary: dict, tasks_root: Path) -> list[dict]:
    rows = []
    for ordinal, result in enumerate(summary["task_summaries"], start=1):
        task_id = result["task_id"]
        task_dir = tasks_root / task_id
        if not task_dir.is_dir():
            raise SystemExit(f"rollout task is missing from tasks/xlsx: {task_id}")
        metadata = load_metadata(task_dir)
        rewards = [float(value) for value in result["rewards"]]
        mean_reward = sum(rewards) / len(rewards)
        rows.append(
            {
                "ordinal": ordinal,
                "task_id": task_id,
                "difficulty_tier": metadata.get("difficulty_tier"),
                "task_family": metadata.get("task_family"),
                "dependency_scope": metadata.get("dependency_scope"),
                "deliverable_types": "+".join(metadata.get("deliverable_types", [])),
                "provenance_id": metadata.get("provenance_id"),
                "leakage_group_id": metadata.get("leakage_group_id"),
                "review_status": "pending_human_review",
                "n_trials": len(rewards),
                "rewards": rewards,
                "mean_reward": round(mean_reward, 6),
                "max_reward": round(max(rewards), 6),
                "min_reward": round(min(rewards), 6),
                "std_dev": round(statistics.pstdev(rewards), 6),
                "passed_trials_ge_0_7": sum(value >= 0.7 for value in rewards),
                "best_trial_ge_0_8": any(value >= 0.8 for value in rewards),
                "task_nonzero": mean_reward > 0,
                "task_perfect": mean_reward == 1,
            }
        )
    return rows


def build_difficulty_summary(rows: list[dict]) -> dict:
    output = {}
    for tier in ("basic", "intermediate", "advanced", "expert"):
        selected = [row for row in rows if row["difficulty_tier"] == tier]
        rewards = [reward for row in selected for reward in row["rewards"]]
        output[tier] = {
            "task_count": len(selected),
            "trial_count": len(rewards),
            "mean_reward": round(sum(rewards) / len(rewards), 6),
            "passed_trials_ge_0_7": sum(value >= 0.7 for value in rewards),
            "trial_pass_rate_ge_0_7": round(
                sum(value >= 0.7 for value in rewards) / len(rewards), 6
            ),
            "zero_mean_task_count": sum(not row["task_nonzero"] for row in selected),
        }
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--summary",
        type=Path,
        default=Path(".coworkevo/runs/xlsx-60-final-complete-summary.json"),
    )
    parser.add_argument("--tasks-root", type=Path, default=Path("tasks/xlsx"))
    parser.add_argument(
        "--output-dir", type=Path, default=Path("datasets/xlsx-rollout-20260901")
    )
    args = parser.parse_args()

    summary = json.loads(args.summary.read_text(encoding="utf-8"))
    rows = build_rows(summary, args.tasks_root)
    if len(rows) != 60 or summary.get("n_tasks") != 60:
        raise SystemExit(f"expected 60 tasks, found {len(rows)}")
    if sum(row["n_trials"] for row in rows) != 180:
        raise SystemExit("expected 180 trials")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    fields = [
        "ordinal",
        "task_id",
        "difficulty_tier",
        "task_family",
        "dependency_scope",
        "deliverable_types",
        "provenance_id",
        "leakage_group_id",
        "review_status",
        "n_trials",
        "rewards",
        "mean_reward",
        "max_reward",
        "min_reward",
        "std_dev",
        "passed_trials_ge_0_7",
        "best_trial_ge_0_8",
        "task_nonzero",
        "task_perfect",
    ]
    csv_path = args.output_dir / "task-results.csv"
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            csv_row = row.copy()
            csv_row["rewards"] = json.dumps(row["rewards"], ensure_ascii=False, separators=(",", ":"))
            writer.writerow(csv_row)

    manifest = {
        "manifest_version": "1.0",
        "status": "pending_human_review",
        "rollout": {
            "harness": summary["harness"],
            "model": summary["model"],
            "generation_time": summary["generation_time"],
            "n_tasks": len(rows),
            "n_runs_per_task": summary["n_runs_per_task"],
            "n_completed_trials": sum(row["n_trials"] for row in rows),
            "source_summary_sha256": sha256(args.summary),
        },
        "metric_definitions": {
            "mean_score": "mean of all 180 trial rewards; each task has three trials",
            "full_task": "task mean_reward == 1.0",
            "high_score_task": "at least one of three trials has reward >= 0.8",
            "nonzero_task": "task mean_reward > 0",
            "zero_task": "task mean_reward == 0",
            "pass_trial": "individual trial reward >= 0.7",
        },
        "reported_accounting": {
            "total_tokens": 120256921,
            "cost_usd": 113.73,
            "source": "rollout cost report supplied with the run; task summary contains no token fields",
        },
        "composition": summary.get("composition", {}),
        "difficulty_summary": build_difficulty_summary(rows),
        "task_results": rows,
    }
    (args.output_dir / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"wrote {csv_path}")
    print(f"wrote {args.output_dir / 'manifest.json'}")


if __name__ == "__main__":
    main()
