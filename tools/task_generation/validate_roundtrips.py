#!/usr/bin/env python3
"""逐题运行参考解法与 Judge，并确认空目录负例会被拒绝。"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[2]
TASKS_ROOT = REPO_ROOT / "tasks"


def _task_roots(domain: str | None) -> list[Path]:
    pattern = f"{domain}/*/instruction.md" if domain else "*/*/instruction.md"
    return sorted(path.parent for path in TASKS_ROOT.glob(pattern) if "templates" not in path.parts)


def _run(command: list[str], root: Path, timeout: int) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["PYTHON_BIN"] = sys.executable
    env["PATH"] = f"{Path(sys.executable).parent}:{env.get('PATH', '')}"
    return subprocess.run(
        command,
        cwd=root,
        env=env,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        timeout=timeout,
        check=False,
    )


def _parse_result(output: str) -> dict[str, Any] | None:
    try:
        payload = json.loads(output.strip())
    except json.JSONDecodeError:
        return None
    return payload if isinstance(payload, dict) else None


def _positive_passed(process: subprocess.CompletedProcess[str]) -> bool:
    if process.returncode != 0:
        return False
    result = _parse_result(process.stdout)
    return result is None or result.get("passed", True) is True


def _negative_rejected(process: subprocess.CompletedProcess[str]) -> bool:
    if process.returncode != 0:
        return True
    result = _parse_result(process.stdout)
    return result is not None and result.get("passed") is False


def _tail(output: str, limit: int = 1200) -> str:
    return output[-limit:].strip()


def validate_task(root: Path, timeout: int) -> list[str]:
    errors: list[str] = []
    with tempfile.TemporaryDirectory(prefix=f"cowork-evo-{root.name}-") as temporary:
        temporary_root = Path(temporary)
        candidate_dir = temporary_root / "candidate"
        candidate_dir.mkdir()
        solve = _run([str(root / "solution/solve.sh"), str(candidate_dir)], root, timeout)
        if solve.returncode != 0:
            errors.append(f"参考解法失败（{solve.returncode}）：{_tail(solve.stdout)}")
            return errors

        positive = _run([str(root / "tests/test.sh"), str(candidate_dir)], root, timeout)
        if not _positive_passed(positive):
            errors.append(f"参考产物未通过 Judge（{positive.returncode}）：{_tail(positive.stdout)}")

        empty_dir = temporary_root / "empty"
        empty_dir.mkdir()
        negative = _run([str(root / "tests/test.sh"), str(empty_dir)], root, timeout)
        if not _negative_rejected(negative):
            errors.append(f"空目录负例没有被拒绝：{_tail(negative.stdout)}")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--domain",
        choices=["ppt", "docx", "xlsx", "cross_office"],
        help="只运行一个主领域",
    )
    parser.add_argument("--timeout", type=int, default=300, help="单个命令超时秒数")
    args = parser.parse_args()
    roots = _task_roots(args.domain)
    if not roots:
        print("没有发现可运行任务。")
        return 1

    failed: list[tuple[Path, list[str]]] = []
    for index, root in enumerate(roots, start=1):
        errors = validate_task(root, args.timeout)
        status = "通过" if not errors else "失败"
        print(f"[{index:02d}/{len(roots):02d}] {status} {root.relative_to(REPO_ROOT)}", flush=True)
        if errors:
            failed.append((root, errors))
    if failed:
        print(f"\n共有 {len(failed)} 个任务未通过端到端验证：")
        for root, errors in failed:
            print(f"- {root.relative_to(REPO_ROOT)}")
            for error in errors:
                print(f"  {error}")
        return 1
    print(f"端到端验证通过：{len(roots)} 个参考解法通过，{len(roots)} 个空目录负例被拒绝。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
