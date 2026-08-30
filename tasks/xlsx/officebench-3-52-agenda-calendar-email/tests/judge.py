#!/usr/bin/env python3
"""独立检查议程到日历/邮件的离线多文件交付。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from email import policy
from email.parser import BytesParser
from pathlib import Path

from openpyxl import load_workbook

INPUT_NAME = "meeting_agenda.xlsx"
EXPECTED_SHA256 = "ffd3161364fde910fbd5491cc1dd19ba346134231e60f217ad668a874af80b0e"
EXPECTED_FILES = ["Bob.ics", "Linda.eml", "Davide.eml", "Charlie.eml"]
EXPECTED_ITEMS = [("Introductions", "Linda", "13:00", "13:30"), ("Agenda", "Davide", "13:30", "14:10"), ("Strategy", "Charlie", "14:10", "14:30")]
TASK_ID = "cowork-evo/officebench-3-52-agenda-calendar-email"
RUBRIC_VERSION = "officebench-3-52-agenda-calendar-email-v0.1"
JUDGE_VERSION = "officebench-3-52-agenda-calendar-email-deterministic-v0.1"
WEIGHTS = {"R001": 0.2, "R002": 0.2, "R003": 0.2, "R004": 0.2, "R005": 0.1, "R006": 0.1}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def criterion(cid: str, passed: bool, evidence: str) -> dict:
    return {"criterion_id": cid, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def parse_ics(path: Path) -> dict[str, str]:
    text = path.read_bytes().decode("utf-8")
    if not text.endswith("\r\n") or text.count("BEGIN:VEVENT\r\n") != 1 or text.count("END:VEVENT\r\n") != 1:
        raise ValueError("iCalendar 边界或 VEVENT 数量异常")
    block = text.split("BEGIN:VEVENT\r\n", 1)[1].split("\r\nEND:VEVENT\r\n", 1)[0]
    props = {}
    for line in block.split("\r\n"):
        if ":" in line:
            key, value = line.split(":", 1)
            props[key.split(";", 1)[0].upper()] = value.replace("\\,", ",").replace("\\;", ";").replace("\\n", "\n").replace("\\\\", "\\")
    required = {"UID", "DTSTAMP", "DTSTART", "DTEND", "SUMMARY", "DESCRIPTION"}
    if not required.issubset(props):
        raise ValueError(f"缺少属性：{sorted(required - set(props))}")
    return props


def parse_email(path: Path):
    message = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
    if not message.get("From") or not message.get("To") or not message.get("Subject") or message.get_content_type() != "text/plain":
        raise ValueError("MIME 头或纯文本正文缺失")
    return message


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    source_path = Path(__file__).parents[1] / "data/input_files" / INPUT_NAME
    results = []
    try:
        actual_files = sorted(path.name for path in args.artifact_dir.iterdir() if path.is_file()) if args.artifact_dir.is_dir() else []
        missing = sorted(set(EXPECTED_FILES) - set(actual_files))
        parse_errors = []
        ics = None
        if (args.artifact_dir / "Bob.ics").is_file():
            try:
                ics = parse_ics(args.artifact_dir / "Bob.ics")
            except Exception as exc:
                parse_errors.append(f"Bob.ics: {exc}")
        emails = {}
        for name in ("Linda", "Davide", "Charlie"):
            path = args.artifact_dir / f"{name}.eml"
            if path.is_file():
                try:
                    emails[name] = parse_email(path)
                except Exception as exc:
                    parse_errors.append(f"{name}.eml: {exc}")
        r001 = not missing and not parse_errors and set(actual_files) == set(EXPECTED_FILES)
        results.append(criterion("R001", r001, f"缺失={missing}; 解析错误={parse_errors}; 文件={actual_files}"))
        timing_ok = bool(ics and ics.get("DTSTART") == "20200501T130000" and ics.get("DTEND") == "20200501T143000")
        results.append(criterion("R002", timing_ok, f"DTSTART={ics.get('DTSTART') if ics else None}; DTEND={ics.get('DTEND') if ics else None}"))
        agenda_text = ics.get("DESCRIPTION", "") if ics else ""
        agenda_ok = all(item in agenda_text and contact in agenda_text for item, contact, _, _ in EXPECTED_ITEMS)
        results.append(criterion("R003", agenda_ok, f"议程正文={agenda_text!r}"))
        email_ok = len(emails) == 3 and all(name.lower() in str(message.get("To", "")).lower() and name in message.get_content() for name, message in emails.items()) and all(item in emails[contact].get_content() and contact in emails[contact].get_content() for item, contact, _, _ in EXPECTED_ITEMS)
        results.append(criterion("R004", email_ok, f"邮件联系人和对应议题={email_ok}"))
        prose_ok = bool(ics and "Meeting agenda" in ics.get("SUMMARY", "") and all("Meeting agenda reminder" in message.get("Subject", "") and "agenda" in message.get_content().lower() for message in emails.values()))
        results.append(criterion("R005", prose_ok, f"日历/邮件提醒语义={prose_ok}"))
        hash_ok = digest(source_path) == EXPECTED_SHA256
        results.append(criterion("R006", hash_ok and set(actual_files) == set(EXPECTED_FILES), f"输入 SHA256={digest(source_path)}"))
        reward = sum(WEIGHTS[item["criterion_id"]] * item["score"] for item in results)
    except Exception as exc:
        results = [criterion("R001", False, f"Judge 读取失败：{exc}")]
        reward = 0.0
    by_id = {item["criterion_id"]: item["score"] for item in results}
    passed = reward >= 0.75 and all(by_id.get(cid) == 1.0 for cid in ("R001", "R002", "R003", "R004"))
    payload = {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": [digest(path) for path in sorted(args.artifact_dir.iterdir()) if path.is_file()] if args.artifact_dir.is_dir() else [], "criterion_results": results, "reward": round(reward, 6), "passed": passed}
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
