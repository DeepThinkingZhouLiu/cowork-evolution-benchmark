#!/usr/bin/env python3
"""使用当前 Judge 和已冻结答案重算 Cross-Office panel，支持断点续跑。"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tomllib
from datetime import datetime
from pathlib import Path


RUNS = {
    "claude-code": "claude-code-sonnet5-cross-office60-20260903",
    "opencode": "opencode-sonnet5-cross-office60-20260903",
    "openhands": "openhands-sonnet5-cross-office60-20260903",
}
TERMINAL_STATUSES = {"completed", "task_failure"}


def _read_json(path: Path) -> dict:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"JSON 根节点不是对象：{path}")
    return payload


def _selected_trial(task: dict) -> Path:
    selected = task.get("selected_attempt")
    attempt = next(
        (item for item in task.get("attempts", []) if item.get("attempt") == selected),
        None,
    )
    if not isinstance(attempt, dict):
        raise ValueError(f"{task.get('task')} 缺少可用的 selected_attempt")
    return Path(str(attempt["trial_dir"])).resolve()


def _task_is_visual(task_dir: Path) -> bool:
    with (task_dir / "task.toml").open("rb") as file:
        payload = tomllib.load(file)
    return "visual" in payload["metadata"]["judge_types"]


def _valid_finished(output_dir: Path) -> bool:
    result_path = output_dir / "combined" / "result.json"
    if not result_path.is_file():
        return False
    try:
        result = _read_json(result_path)
    except Exception:
        return False
    return result.get("status") in TERMINAL_STATUSES and isinstance(
        result.get("combined_reward"), (int, float)
    )


def _run_deterministic(
    *,
    python: Path,
    task_dir: Path,
    candidate_dir: Path,
    destination: Path,
) -> Path:
    destination.mkdir(parents=True, exist_ok=True)
    result_path = destination / "result.json"
    reward_path = destination / "reward.txt"
    environment = os.environ.copy()
    environment.update(
        {
            "PYTHON_BIN": str(python),
            "ARTIFACT_DIR": str(candidate_dir),
            "INPUT_DIR": str(task_dir / "data" / "input_files"),
            "JUDGE_RESULT_PATH": str(result_path),
            "RESULT_FILE": str(result_path),
            "REWARD_PATH": str(reward_path),
            "REWARD_FILE": str(reward_path),
            "SCORING_VERSION": "v2",
            "PYTHONDONTWRITEBYTECODE": "1",
        }
    )
    completed = subprocess.run(
        ["bash", str(task_dir / "tests" / "test.sh"), str(candidate_dir)],
        cwd=task_dir.parents[1],
        env=environment,
        capture_output=True,
        text=True,
        timeout=1800,
        check=False,
    )
    (destination / "stdout.txt").write_text(completed.stdout, encoding="utf-8")
    (destination / "stderr.txt").write_text(completed.stderr, encoding="utf-8")
    if not result_path.is_file():
        raise RuntimeError(
            f"确定性 Judge 未生成 result.json（exit={completed.returncode}）"
        )
    result = _read_json(result_path)
    reward = result.get("deterministic_reward", result.get("reward", result.get("score")))
    if isinstance(reward, bool) or not isinstance(reward, (int, float)) or not 0 <= float(reward) <= 1:
        raise RuntimeError("确定性 Judge 的 reward 无效")
    return result_path


def _run_composite(
    *,
    repo: Path,
    python: Path,
    codex: Path,
    config: Path,
    env_file: Path,
    task_dir: Path,
    candidate_dir: Path,
    deterministic_result: Path,
    output_dir: Path,
) -> int:
    environment = os.environ.copy()
    environment["PYTHONPATH"] = str(repo / "src")
    environment["COWORK_EVO_CODEX_BINARY"] = str(codex)
    command = [
        str(python),
        "-m",
        "cowork_evo_bench.cli",
        "judge-visual",
        "--task-dir",
        str(task_dir),
        "--candidate-dir",
        str(candidate_dir),
        "--deterministic-result",
        str(deterministic_result),
        "--config",
        str(config),
        "--env-file",
        str(env_file),
        "--output-dir",
        str(output_dir),
    ]
    completed = subprocess.run(
        command,
        cwd=repo,
        env=environment,
        capture_output=True,
        text=True,
        timeout=2400,
        check=False,
    )
    log_dir = output_dir.parent / f".{output_dir.name}-orchestration"
    log_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / "stdout.txt").write_text(completed.stdout, encoding="utf-8")
    (log_dir / "stderr.txt").write_text(completed.stderr, encoding="utf-8")
    return completed.returncode


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument(
        "--runs-root",
        type=Path,
        default=Path("REMOVED_LOCAL_PATH"),
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=Path(".coworkevo/empirical-calibration/cross-office-hybrid-20260904"),
    )
    parser.add_argument(
        "--python",
        type=Path,
        default=Path("REMOVED_LOCAL_PATH"),
    )
    parser.add_argument("--codex", type=Path, default=Path("REMOVED_LOCAL_PATH"))
    parser.add_argument("--config", type=Path, default=Path("configs/judging/codex-vlm.json"))
    parser.add_argument("--env-file", type=Path, default=Path(".env"))
    parser.add_argument("--harness", choices=["all", *RUNS], default="all")
    parser.add_argument("--task")
    arguments = parser.parse_args()

    repo = arguments.repo.resolve()
    output_root = arguments.output_root
    if not output_root.is_absolute():
        output_root = (repo / output_root).resolve()
    config = arguments.config if arguments.config.is_absolute() else (repo / arguments.config).resolve()
    env_file = arguments.env_file if arguments.env_file.is_absolute() else (repo / arguments.env_file).resolve()
    harnesses = list(RUNS) if arguments.harness == "all" else [arguments.harness]
    progress_path = output_root / "progress.jsonl"
    progress_path.parent.mkdir(parents=True, exist_ok=True)

    for harness in harnesses:
        source = arguments.runs_root / RUNS[harness]
        manifest = _read_json(source / "results.json")
        task_rows = manifest.get("tasks", [])
        if len(task_rows) != 60 or len({row.get("task") for row in task_rows}) != 60:
            raise RuntimeError(f"{harness} 的冻结答案清单不是唯一 60 题")
        for index, row in enumerate(task_rows, 1):
            task_name = str(row["task"])
            if arguments.task and task_name != arguments.task:
                continue
            task_dir = (repo / "tasks" / "cross_office" / task_name).resolve()
            trial_dir = _selected_trial(row)
            candidate_dir = trial_dir / "artifacts" / "app" / "output"
            final_dir = output_root / "panel" / harness / task_name
            event = {
                "timestamp": datetime.now().astimezone().isoformat(),
                "harness": harness,
                "task": task_name,
                "index": index,
                "visual": _task_is_visual(task_dir),
                "candidate_dir": str(candidate_dir),
            }
            if _valid_finished(final_dir):
                event["action"] = "skip_completed"
            elif final_dir.exists():
                event["action"] = "skip_existing_nonterminal"
                event["error"] = "已存在非终态评分目录，为保留审计证据未覆盖"
            else:
                deterministic_dir = output_root / "deterministic-source" / harness / task_name
                try:
                    deterministic = _run_deterministic(
                        python=arguments.python,
                        task_dir=task_dir,
                        candidate_dir=candidate_dir,
                        destination=deterministic_dir,
                    )
                    code = _run_composite(
                        repo=repo,
                        python=arguments.python,
                        codex=arguments.codex,
                        config=config,
                        env_file=env_file,
                        task_dir=task_dir,
                        candidate_dir=candidate_dir,
                        deterministic_result=deterministic,
                        output_dir=final_dir,
                    )
                    result = _read_json(final_dir / "combined" / "result.json")
                    event.update(
                        {
                            "action": "scored",
                            "cli_exit": code,
                            "status": result.get("status"),
                            "deterministic_reward": result.get("deterministic_reward"),
                            "visual_reward": result.get("visual_reward"),
                            "combined_reward": result.get("combined_reward"),
                        }
                    )
                except Exception as error:
                    event["action"] = "orchestration_failure"
                    event["error"] = f"{type(error).__name__}: {error}"
            with progress_path.open("a", encoding="utf-8") as stream:
                stream.write(json.dumps(event, ensure_ascii=False) + "\n")
            print(json.dumps(event, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
