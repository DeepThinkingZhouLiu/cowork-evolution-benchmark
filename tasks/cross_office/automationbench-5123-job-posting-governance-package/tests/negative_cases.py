#!/usr/bin/env python3
"""AutomationBench 5123 的任务专属 v2 分层负例。"""

from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Callable

ROOT = Path(__file__).resolve().parents[1]
GOLDEN = ROOT / "solution" / "Golden"
INPUTS = ROOT / "environment" / "input_files"
OUTPUTS = ("Job_Posting_Audit.xlsx", "Job_Posting_Control_Brief.docx", "Job_Posting_Review.pptx", "Job_Posting_Email_Drafts.mbox", "Job_Posting_Action_Drafts.json", "Job_Posting_Slack_Control.json", "Job_Posting_Execution_Control.json")
XLSX, DOCX, PPTX, MBOX, ACTIONS, SLACK, EXECUTION = OUTPUTS
Mutator = Callable[[Path], None]


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法载入{path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


NEG = _load("job_posting_5123_negative_v1_helpers", ROOT / "tests" / "negative_cases_v1.py")


def _input_only(root: Path) -> None:
    for path in list(root.iterdir()):
        if path.is_file():
            path.unlink()
    for source in INPUTS.iterdir():
        if source.is_file():
            shutil.copy2(source, root / source.name)


def _edit(sheet: str, cell: str, value, *, remove_formula: bool = False) -> Mutator:
    return lambda root: NEG.edit_cell(root / XLSX, sheet, cell, value=value, remove_formula=remove_formula)


def _json(name: str, mutator: Callable[[dict], None]) -> Mutator:
    return lambda root: NEG.mutate_json(root / name, mutator)


def _message(index: int, mutator: Callable[[dict], None]) -> Mutator:
    return lambda root: NEG.mutate_message(root / MBOX, index, mutator)


def _word_account_age(root: Path) -> None:
    NEG.replace_docx_text(root / DOCX, "90", "91", row_token="Account Executive", min_cells=11)


def _ppt_senior_route(root: Path) -> None:
    NEG.replace_pptx_text(root / PPTX, "vp-people@company.example.com", "recruiting@company.example.com", row_token="Senior Backend Engineer", min_cells=10)


def _cases() -> list[tuple[str, Mutator, str]]:
    cases: list[tuple[str, Mutator, str]] = []
    for name in OUTPUTS:
        cases.append((f"missing-{name}", lambda root, item=name: (root / item).unlink(), "fatal_zero"))
        cases.append((f"corrupt-{name}", lambda root, item=name: (root / item).write_bytes(b"not-a-valid-artifact"), "fatal_zero"))
    cases.extend([
        ("mbox-empty", lambda root: (root / MBOX).write_bytes(b""), "fatal_zero"),
        ("unmodified-input-only", _input_only, "fatal_zero"),
        ("source-account-original-date-local", _edit("Source_Positions", "F4", 46036), "partial"),
        ("formula-account-stale-boundary-local", _edit("Posting_Decisions", "N4", "YES", remove_formula=True), "partial"),
        ("direct-account-recipient-local", _edit("Direct_Repost_Queue", "I3", "wrong@company.example.com"), "partial"),
        ("stale-senior-recipient-local", _edit("Stale_Approval_Queue", "K2", "recruiting@company.example.com"), "partial"),
        ("no-action-marketing-decision-local", _edit("No_Action_Control", "G2", "DIRECT_REPOST"), "partial"),
        ("mail-frontend-subject-local", _message(0, lambda state: state.update({"subject": "Updated role: Frontend Engineer"})), "partial"),
        ("mail-account-wrong-route", _message(1, lambda state: state.update({"to": "vp-people@company.example.com"})), "partial"),
        ("mail-marketing-forbidden-scope", _message(0, lambda state: state.update({"body": state["body"] + "\nMarketing Manager"})), "partial"),
        ("actions-account-age-local", _json(ACTIONS, lambda payload: payload["decisions"][2].update({"age_days": 91})), "partial"),
        ("actions-marketing-direct-repost", _json(ACTIONS, lambda payload: payload["decisions"][1].update({"decision": "DIRECT_REPOST"})), "partial"),
        ("actions-candidate-outreach-executed", _json(ACTIONS, lambda payload: payload["decisions"][0].update({"candidate_outreach_executed": True})), "partial"),
        ("slack-posted", _json(SLACK, lambda payload: payload.update({"posted": True, "posted_count": 1})), "partial"),
        ("execution-emails-sent", _json(EXECUTION, lambda payload: payload["email"].update({"sent_count": 6, "state": "SENT"})), "partial"),
        ("word-account-age-local", _word_account_age, "partial"),
        ("ppt-senior-route-local", _ppt_senior_route, "partial"),
        ("extra-file", lambda root: (root / "unexpected.txt").write_text("unexpected", encoding="utf-8"), "partial"),
    ])
    return cases


def main() -> None:
    parser = argparse.ArgumentParser(); parser.add_argument("--report", type=Path); args = parser.parse_args()
    judge = _load("job_posting_5123_v2_judge", ROOT / "tests" / "judge.py")
    judge.V1.render = lambda path: ({".xlsx":11, ".docx":3, ".pptx":3}[path.suffix.casefold()], "render-ok")
    if judge.evaluate(GOLDEN).get("reward") != 1.0:
        raise AssertionError("Golden v2必须为1.0")
    results = []
    for name, mutator, layer in _cases():
        with tempfile.TemporaryDirectory(prefix=f"job-posting-5123-v2-{name}-") as temporary:
            output = Path(temporary) / "output"; shutil.copytree(GOLDEN, output); mutator(output); result = judge.evaluate(output); reward = float(result.get("reward", 0))
            if layer == "fatal_zero" and reward != 0:
                raise AssertionError(f"{name}应硬归零，实际{reward}")
            if layer == "partial" and not (0 < reward < 1):
                raise AssertionError(f"{name}应保留连续部分分，实际{reward}")
            results.append({"case":name,"layer":layer,"reward":reward,"passed":bool(result.get("passed")),"penalties":result.get("penalties",[])})
    partial = [item["reward"] for item in results if item["layer"] == "partial"]
    report = {"total":len(results),"fatal_zero":sum(item["layer"] == "fatal_zero" for item in results),"partial_continuous":len(partial),"partial_reward_range":[min(partial),max(partial)],"cases":results}
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True); args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()
