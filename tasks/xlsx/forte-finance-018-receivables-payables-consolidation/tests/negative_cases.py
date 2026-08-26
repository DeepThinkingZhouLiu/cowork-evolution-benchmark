#!/usr/bin/env python3
"""Golden、fresh replay 与 7 个针对性错误交付。"""

from __future__ import annotations

import csv
import importlib.util
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True

from openpyxl import load_workbook


TASK_ROOT = Path(__file__).resolve().parents[1]
SOLUTION = TASK_ROOT / "solution"
INPUT_DIR = TASK_ROOT / "data" / "input_files"
OUTPUT_NAMES = [
    "reconcile_workbooks.py",
    "未付统计.csv",
    "未收统计.csv",
    "僵尸账款核查.txt",
    "往来款跨期整合.xlsx",
]

spec = importlib.util.spec_from_file_location("forte_finance_018_judge", Path(__file__).with_name("judge.py"))
if spec is None or spec.loader is None:
    raise RuntimeError("无法加载 Judge")
judge = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = judge
spec.loader.exec_module(judge)


def copy_candidate(target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    for name in OUTPUT_NAMES:
        shutil.copy2(SOLUTION / name, target / name)


def scores(result: dict) -> dict[str, float]:
    return {item["criterion_id"]: item["score"] for item in result["criterion_results"]}


def assert_full(label: str, path: Path) -> None:
    result = judge.evaluate(path, render=False)
    assert result["passed"] and result["reward"] == 1.0, (label, result["reward"], result["hurdle_failures"])


def assert_rejected(label: str, path: Path, criterion_id: str) -> None:
    result = judge.evaluate(path, render=False)
    result_scores = scores(result)
    assert not result["passed"] and result["reward"] < 0.8, (label, result["reward"])
    assert result_scores[criterion_id] == 0.0, (label, criterion_id, result_scores[criterion_id])


def read_gb_csv(path: Path) -> list[list[str]]:
    return list(csv.reader(path.read_text(encoding="gb18030").splitlines()))


def write_gb_csv(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", encoding="gb18030", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\r\n")
        writer.writerows(rows)


def case_golden(base: Path) -> None:
    candidate = base / "golden"
    copy_candidate(candidate)
    assert_full("golden", candidate)


def case_fresh(base: Path) -> None:
    candidate = base / "fresh"
    subprocess.run(
        [
            sys.executable,
            str(SOLUTION / "reconcile_workbooks.py"),
            "--input-dir",
            str(INPUT_DIR),
            "--output-dir",
            str(candidate),
        ],
        check=True,
        timeout=120,
    )
    assert_full("fresh", candidate)


def case_missing(base: Path) -> None:
    candidate = base / "missing"
    copy_candidate(candidate)
    (candidate / "往来款跨期整合.xlsx").unlink()
    assert_rejected("missing workbook", candidate, "R001")


def case_wrong_amount(base: Path) -> None:
    candidate = base / "wrong-amount"
    copy_candidate(candidate)
    path = candidate / "未付统计.csv"
    rows = read_gb_csv(path)
    rows[1][2] = "20,001.00"
    write_gb_csv(path, rows)
    assert_rejected("wrong unpaid amount", candidate, "R002")


def case_omitted_row(base: Path) -> None:
    candidate = base / "omitted-row"
    copy_candidate(candidate)
    path = candidate / "未收统计.csv"
    rows = read_gb_csv(path)
    write_gb_csv(path, rows[:-1])
    assert_rejected("omitted unreceived row", candidate, "R003")


def case_fake_zombie(base: Path) -> None:
    candidate = base / "fake-zombie"
    copy_candidate(candidate)
    path = candidate / "僵尸账款核查.txt"
    path.write_text("僵尸账款核查结论：存在僵尸账款。\n- 【虚构客商】 | 应收账款 | 99,999.00\n", encoding="utf-8")
    assert_rejected("fake zombie", candidate, "R004")


def case_hardcoded_processor(base: Path) -> None:
    candidate = base / "hardcoded"
    copy_candidate(candidate)
    path = candidate / "reconcile_workbooks.py"
    path.write_text(
        """#!/usr/bin/env python3
import argparse
import shutil
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument('--input-dir', required=True)
parser.add_argument('--output-dir', required=True, type=Path)
args = parser.parse_args()
source = Path(__file__).resolve().parent
args.output_dir.mkdir(parents=True, exist_ok=True)
for name in ['reconcile_workbooks.py', '未付统计.csv', '未收统计.csv', '僵尸账款核查.txt', '往来款跨期整合.xlsx']:
    destination = args.output_dir / name
    if (source / name).resolve() != destination.resolve():
        shutil.copy2(source / name, destination)
""",
        encoding="utf-8",
    )
    assert_rejected("hardcoded processor", candidate, "R006")


def case_cross_artifact_conflict(base: Path) -> None:
    candidate = base / "cross-artifact-conflict"
    copy_candidate(candidate)
    path = candidate / "往来款跨期整合.xlsx"
    workbook = load_workbook(path)
    workbook["未付统计"]["C2"] = 999999
    workbook.save(path)
    workbook.close()
    assert_rejected("cross artifact conflict", candidate, "R007")


def case_unsupported_assertion(base: Path) -> None:
    candidate = base / "unsupported-assertion"
    copy_candidate(candidate)
    path = candidate / "僵尸账款核查.txt"
    text = path.read_text(encoding="utf-8")
    path.write_text(text + "- 【无来源客商】 | 其他应收款 | 88,888.00\n", encoding="utf-8")
    assert_rejected("unsupported assertion", candidate, "R004")


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="cowork-evo-forte-negative-") as directory:
        base = Path(directory)
        case_golden(base)
        case_fresh(base)
        case_missing(base)
        case_wrong_amount(base)
        case_omitted_row(base)
        case_fake_zombie(base)
        case_hardcoded_processor(base)
        case_cross_artifact_conflict(base)
        case_unsupported_assertion(base)
    print("PASS: Golden=1.0, fresh replay=1.0, 7/7 negative cases rejected")


if __name__ == "__main__":
    main()
