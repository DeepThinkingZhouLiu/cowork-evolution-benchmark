#!/usr/bin/env python3
"""使用 Harbor 和容器内 Claude Code 顺序运行 60 道 PPT 任务。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_HARBOR = Path(os.environ.get("COWORK_EVO_HARBOR_BINARY", "harbor"))
DEFAULT_SETTINGS = Path.home() / ".claude" / "settings.json"
DEFAULT_BINARY = (
    Path.home()
    / ".local/lib/node_modules/@anthropic-ai/claude-code/bin/claude.exe"
)
AGENT_IMPORT = (
    "cowork_evo_bench.adapters.host_binary_claude_code:HostBinaryClaudeCode"
)
FIRST_TASK = "pptarena-001-kazakh-to-english-preserve-french"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def ordered_tasks(tasks_root: Path) -> list[Path]:
    tasks = sorted(
        path
        for path in tasks_root.iterdir()
        if path.is_dir() and (path / "task.toml").is_file()
    )
    by_name = {path.name: path for path in tasks}
    if len(tasks) != 60:
        raise RuntimeError(f"预期 60 道 PPT 任务，实际发现 {len(tasks)} 道。")
    if FIRST_TASK not in by_name:
        raise RuntimeError(f"找不到首个 Smoke 任务：{FIRST_TASK}")
    return [by_name[FIRST_TASK], *[p for p in tasks if p.name != FIRST_TASK]]


def completed_trial(job_dir: Path) -> dict[str, Any] | None:
    results = sorted(job_dir.glob("*/result.json"))
    if len(results) != 1:
        return None
    try:
        payload = json.loads(results[0].read_text())
    except (OSError, json.JSONDecodeError):
        return None
    if not payload.get("finished_at"):
        return None
    return payload


def compact_result(index: int, task: Path, job_dir: Path) -> dict[str, Any]:
    payload = completed_trial(job_dir)
    if payload is None:
        return {
            "index": index,
            "task": task.name,
            "job_dir": str(job_dir),
            "completed": False,
        }

    rewards = ((payload.get("verifier_result") or {}).get("rewards") or {})
    agent_result = payload.get("agent_result") or {}
    exception = payload.get("exception_info")
    return {
        "index": index,
        "task": task.name,
        "job_dir": str(job_dir),
        "trial_name": payload.get("trial_name"),
        "completed": True,
        "reward": rewards.get("reward"),
        "exception": exception,
        "input_tokens": agent_result.get("n_input_tokens"),
        "cache_tokens": agent_result.get("n_cache_tokens"),
        "output_tokens": agent_result.get("n_output_tokens"),
        "cost_usd": agent_result.get("cost_usd"),
        "started_at": payload.get("started_at"),
        "finished_at": payload.get("finished_at"),
    }


def write_manifest(
    output_root: Path,
    *,
    settings: Path,
    binary: Path,
    model: str,
    cli_version: str,
) -> None:
    manifest = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "git_commit": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=REPO_ROOT, text=True
        ).strip(),
        "agent": AGENT_IMPORT,
        "model": model,
        "claude_code_version": cli_version,
        "settings_path": str(settings),
        "settings_sha256": sha256(settings),
        "host_binary": str(binary),
        "host_binary_sha256": sha256(binary),
        "provider_host": "api.zcloudapi.com",
        "n_concurrent": 1,
        "n_attempts": 1,
        "max_turns": 80,
        "reasoning_effort": "xhigh",
        "disallowed_tools": ["WebFetch", "WebSearch"],
        "task_count": 60,
    }
    (output_root / "run-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output-root",
        type=Path,
        default=REPO_ROOT / "jobs/claude-sonnet5-ppt60-20260830",
    )
    parser.add_argument("--harbor", type=Path, default=DEFAULT_HARBOR)
    parser.add_argument("--settings", type=Path, default=DEFAULT_SETTINGS)
    parser.add_argument("--host-binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--model", default="claude-sonnet-5")
    parser.add_argument("--cli-version", default="2.1.251")
    parser.add_argument("--start-index", type=int, default=1)
    parser.add_argument("--end-index", type=int, default=60)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not 1 <= args.start_index <= args.end_index <= 60:
        raise ValueError("--start-index/--end-index 必须满足 1 <= start <= end <= 60。")
    for required in (args.harbor, args.settings, args.host_binary):
        if not required.is_file():
            raise FileNotFoundError(required)

    tasks = ordered_tasks(REPO_ROOT / "tasks/ppt")
    output_root = args.output_root.resolve()
    log_root = output_root / "runner-logs"
    log_root.mkdir(parents=True, exist_ok=True)
    write_manifest(
        output_root,
        settings=args.settings,
        binary=args.host_binary,
        model=args.model,
        cli_version=args.cli_version,
    )

    progress_path = output_root / "progress.jsonl"
    environment = os.environ.copy()
    existing_pythonpath = environment.get("PYTHONPATH")
    environment["PYTHONPATH"] = str(REPO_ROOT / "src") + (
        os.pathsep + existing_pythonpath if existing_pythonpath else ""
    )

    for index, task in enumerate(tasks, start=1):
        if index < args.start_index or index > args.end_index:
            continue
        job_name = f"{index:02d}-{task.name}"
        job_dir = output_root / job_name
        result = completed_trial(job_dir)
        if result is not None:
            summary = compact_result(index, task, job_dir)
            print(
                f"SKIP {index:02d}/60 {task.name} reward={summary.get('reward')}",
                flush=True,
            )
            continue

        print(f"START {index:02d}/60 {task.name}", flush=True)
        command = [
            str(args.harbor),
            "run",
            "--path",
            str(task),
            "--agent",
            AGENT_IMPORT,
            "--model",
            args.model,
            "--agent-kwarg",
            f"version={args.cli_version}",
            "--agent-kwarg",
            f"config={args.settings}",
            "--agent-kwarg",
            f"host_binary={args.host_binary}",
            "--agent-kwarg",
            "provider_host=api.zcloudapi.com",
            "--agent-kwarg",
            "max_turns=80",
            "--agent-kwarg",
            "reasoning_effort=xhigh",
            "--agent-kwarg",
            "disallowed_tools=WebFetch,WebSearch",
            "--allow-agent-host",
            "api.zcloudapi.com",
            "--n-concurrent",
            "1",
            "--max-retries",
            "0",
            "--jobs-dir",
            str(output_root),
            "--job-name",
            job_name,
            "--quiet",
        ]
        log_path = log_root / f"{job_name}.log"
        with log_path.open("w") as log_file:
            process = subprocess.run(
                command,
                cwd=REPO_ROOT,
                env=environment,
                stdout=log_file,
                stderr=subprocess.STDOUT,
                text=True,
                check=False,
            )

        summary = compact_result(index, task, job_dir)
        summary["harbor_exit_code"] = process.returncode
        with progress_path.open("a") as progress_file:
            progress_file.write(json.dumps(summary, ensure_ascii=False) + "\n")

        print(
            f"DONE  {index:02d}/60 {task.name} "
            f"reward={summary.get('reward')} "
            f"exception={bool(summary.get('exception'))} "
            f"exit={process.returncode}",
            flush=True,
        )

    print(f"COMPLETE output_root={output_root}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
