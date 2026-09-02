#!/usr/bin/env python3
"""OfficeBench 3-6 团队培训协调包的任务专属 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import mailbox
import re
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time
from email import message_from_binary_file, policy
from email.utils import parseaddr
from pathlib import Path
from typing import Any

from docx import Document
from openpyxl import load_workbook


SOURCE_NAME = "team_training_schedule.xlsx"
SOURCE_SHA256 = "a5628f45d9f48eccd8424634fda4d3c9048200571cc03db1e6d34cf9447498f4"
XLSX_NAME = "Team_Training_Coordination_Audit.xlsx"
DOCX_NAME = "Team_Training_Coordination_Brief.docx"
ICS_NAME = "Team_Training_Master.ics"
MBOX_NAME = "Team_Training_Reminders.mbox"
REQUIRED_FILES = (XLSX_NAME, DOCX_NAME, ICS_NAME, MBOX_NAME)
TRAINING_DATE = date(2020, 5, 8)

WEIGHTS = {
    "R001": 5,
    "R002": 9,
    "R003": 8,
    "R004": 8,
    "R005": 9,
    "R006": 8,
    "R007": 9,
    "R008": 8,
    "R009": 8,
    "R010": 8,
    "R011": 6,
    "R012": 5,
    "R013": 4,
    "R014": 5,
}

HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.25,
    "R003": 0.30,
    "R004": 0.35,
    "R005": 0.50,
    "R006": 0.35,
    "R007": 0.30,
    "R008": 0.35,
    "R009": 0.40,
    "R010": 0.45,
    "R011": 0.40,
    "R012": 0.40,
}

EXPECTED_ROWS = (
    (2, "Alice", "Section A", "9:00 AM - 11:00 AM"),
    (3, "Bob", "Section B", "11:00 AM - 1:00 PM"),
    (4, "Carol", "Section C", "2:00 PM - 4:00 PM"),
    (5, "David", "Section A", "9:00 AM - 11:00 AM"),
    (6, "Eva", "Section B", "11:00 AM - 1:00 PM"),
    (7, "Frank", "Section C", "2:00 PM - 4:00 PM"),
    (8, "Grace", "Section A", "9:00 AM - 11:00 AM"),
    (9, "Hank", "Section B", "11:00 AM - 1:00 PM"),
    (10, "Ivy", "Section C", "2:00 PM - 4:00 PM"),
    (11, "Jack", "Section A", "9:00 AM - 11:00 AM"),
    (12, "Kelly", "Section B", "11:00 AM - 1:00 PM"),
    (13, "Leo", "Section C", "2:00 PM - 4:00 PM"),
)

SESSION_TIMES = {
    "Section A": (time(9, 0), time(11, 0)),
    "Section B": (time(11, 0), time(13, 0)),
    "Section C": (time(14, 0), time(16, 0)),
}


@dataclass(frozen=True)
class SourceMember:
    source_row: int
    name: str
    section: str
    source_time: str
    start: time
    end: time


class Checks:
    def __init__(self) -> None:
        self.values: dict[str, dict[str, Any]] = {
            criterion: {"passed": True, "details": []} for criterion in WEIGHTS
        }

    def fail(self, criterion: str, detail: str) -> None:
        self.values[criterion]["passed"] = False
        if detail not in self.values[criterion]["details"]:
            self.values[criterion]["details"].append(detail)

    def passed(self, criterion: str) -> bool:
        return bool(self.values[criterion]["passed"])


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def source_directory(task_root: Path) -> Path:
    harbor_path = Path("/app/input_files")
    if harbor_path.is_dir():
        return harbor_path
    return task_root / "data" / "input_files"


def parse_clock(value: str) -> time:
    match = re.fullmatch(r"(\d{1,2}):(\d{2}) (AM|PM)", value.strip())
    if not match:
        raise ValueError(value)
    hour = int(match.group(1)) % 12
    if match.group(3) == "PM":
        hour += 12
    return time(hour, int(match.group(2)))


def read_source_truth(task_root: Path) -> list[SourceMember]:
    source = source_directory(task_root) / SOURCE_NAME
    if not source.is_file() or sha256(source) != SOURCE_SHA256:
        raise ValueError("冻结源表缺失或 SHA256 错误")
    workbook = load_workbook(source, read_only=True, data_only=False)
    if workbook.sheetnames != ["Training Schedule"]:
        raise ValueError("源表工作表结构错误")
    sheet = workbook["Training Schedule"]
    if tuple(sheet.cell(1, column).value for column in range(1, 4)) != ("Name", "Section", "Time"):
        raise ValueError("源表表头错误")
    observed: list[tuple[int, str, str, str]] = []
    members: list[SourceMember] = []
    for row in range(2, 14):
        name, section, source_time = (str(sheet.cell(row, column).value) for column in range(1, 4))
        observed.append((row, name, section, source_time))
        pieces = source_time.split(" - ")
        if len(pieces) != 2:
            raise ValueError(f"源表第 {row} 行时间错误")
        members.append(SourceMember(row, name, section, source_time, parse_clock(pieces[0]), parse_clock(pieces[1])))
    if tuple(observed) != EXPECTED_ROWS:
        raise ValueError("源表 12 行人员真值与冻结版本不一致")
    return members


def grouped_members(members: list[SourceMember]) -> dict[str, list[SourceMember]]:
    grouped = {section: [] for section in SESSION_TIMES}
    for member in members:
        grouped[member.section].append(member)
    return grouped


def section_slug(section: str) -> str:
    return section.lower().replace(" ", "-")


def section_uid(section: str) -> str:
    return f"{section_slug(section)}-training-20200508@evocowork.invalid"


def time_value(value: Any) -> time | None:
    if isinstance(value, time):
        return value
    if isinstance(value, datetime):
        return value.time()
    return None


def date_value(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def formula_contains(value: Any, *parts: str) -> bool:
    if not isinstance(value, str) or not value.startswith("="):
        return False
    normalized = re.sub(r"[\s'$]", "", value).upper()
    return all(re.sub(r"[\s'$]", "", part).upper() in normalized for part in parts)


def inspect_ooxml(path: Path, checks: Checks) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            dangerous = [
                name
                for name in names
                if "vbaproject" in name.lower()
                or name.startswith("xl/externalLinks/")
                or "/embeddings/" in name
                or "oleobject" in name.lower()
            ]
            if dangerous:
                checks.fail("R005", f"OOXML 包含不允许的外链/宏/嵌入：{dangerous}")
                checks.fail("R011", "Office 交付包含危险或无来源对象")
    except Exception as exc:
        checks.fail("R001", f"OOXML 无法解析：{exc}")


def check_workbook(path: Path, members: list[SourceMember], checks: Checks) -> dict[str, Any]:
    context: dict[str, Any] = {}
    try:
        inspect_ooxml(path, checks)
        formulas = load_workbook(path, data_only=False)
        values = load_workbook(path, data_only=True)
    except Exception as exc:
        for criterion in ("R002", "R003", "R004", "R005", "R012", "R014"):
            checks.fail(criterion, f"Excel 无法打开：{exc}")
        return context

    expected_sheets = ["Roster_Control", "Session_Control", "Reminder_Control", "Delivery_Summary"]
    if formulas.sheetnames != expected_sheets:
        checks.fail("R005", f"Excel 工作表应为 {expected_sheets}，实际为 {formulas.sheetnames}")
        return context

    roster_f, roster_v = formulas["Roster_Control"], values["Roster_Control"]
    roster_headers = [roster_f.cell(1, column).value for column in range(1, 12)]
    expected_roster_headers = [
        "Source Row", "Name", "Section", "Source Time", "Start", "End", "Duration (Hours)",
        "Training Date", "Reserved Address", "Delivery Status", "Session UID",
    ]
    if roster_headers != expected_roster_headers or roster_f.max_row != 13 or roster_f.max_column != 11:
        checks.fail("R002", "Roster_Control 结构、行数或表头不正确")
    for output_row, member in enumerate(members, start=2):
        observed = [roster_f.cell(output_row, column).value for column in range(1, 7)]
        expected = [member.source_row, member.name, member.section, member.source_time, member.start, member.end]
        if observed[:4] != expected[:4] or time_value(observed[4]) != member.start or time_value(observed[5]) != member.end:
            checks.fail("R002", f"Roster_Control 第 {output_row} 行与源表不一致")
        if date_value(roster_f.cell(output_row, 8).value) != TRAINING_DATE:
            checks.fail("R003", f"Roster_Control 第 {output_row} 行培训日期错误")
        expected_address = f"{member.name.lower()}@example.invalid"
        if roster_v.cell(output_row, 7).value != 2:
            checks.fail("R004", f"Roster_Control 第 {output_row} 行时长缓存不为 2")
        if roster_v.cell(output_row, 9).value != expected_address:
            checks.fail("R011", f"Roster_Control 第 {output_row} 行保留地址错误")
        if roster_f.cell(output_row, 10).value != "BLOCKED_MISSING_REAL_EMAIL":
            checks.fail("R011", f"Roster_Control 第 {output_row} 行未标记阻塞")
        if roster_v.cell(output_row, 11).value != section_uid(member.section):
            checks.fail("R014", f"Roster_Control 第 {output_row} 行 UID 不可追溯")
        if not formula_contains(roster_f.cell(output_row, 7).value, "ROUND", "MOD", f"F{output_row}-E{output_row}"):
            checks.fail("R004", f"Roster_Control!G{output_row} 不是时长公式")
        if not formula_contains(roster_f.cell(output_row, 9).value, "LOWER", f"B{output_row}", "EXAMPLE.INVALID"):
            checks.fail("R004", f"Roster_Control!I{output_row} 不是地址公式")

    grouped = grouped_members(members)
    sessions_f, sessions_v = formulas["Session_Control"], values["Session_Control"]
    expected_session_headers = [
        "Section", "Training Date", "Start", "End", "Duration (Hours)", "Member Count", "Members",
        "Previous End", "Gap (Minutes)", "Overlap Flag", "Session UID", "Calendar Status",
    ]
    if [sessions_f.cell(1, column).value for column in range(1, 13)] != expected_session_headers:
        checks.fail("R003", "Session_Control 表头不正确")
    expected_names: dict[str, str] = {
        section: " | ".join(member.name for member in grouped[section]) for section in SESSION_TIMES
    }
    session_context: dict[str, dict[str, Any]] = {}
    for row, (section, (start, end)) in enumerate(SESSION_TIMES.items(), start=2):
        observed_section = sessions_f.cell(row, 1).value
        if observed_section != section:
            checks.fail("R003", f"Session_Control 第 {row} 行 Section 错误")
        if date_value(sessions_f.cell(row, 2).value) != TRAINING_DATE:
            checks.fail("R003", f"{section} 日期错误")
        if time_value(sessions_f.cell(row, 3).value) != start or time_value(sessions_f.cell(row, 4).value) != end:
            checks.fail("R003", f"{section} 开始或结束时间错误")
        if sessions_f.cell(row, 7).value != expected_names[section]:
            checks.fail("R003", f"{section} 人员名单错误")
        if sessions_f.cell(row, 11).value != section_uid(section) or sessions_f.cell(row, 12).value != "TENTATIVE":
            checks.fail("R003", f"{section} UID 或状态错误")
        if sessions_v.cell(row, 5).value != 2 or sessions_v.cell(row, 6).value != 4:
            checks.fail("R004", f"{section} 时长或人数公式缓存错误")
        expected_gap = {2: None, 3: 0, 4: 60}[row]
        if sessions_v.cell(row, 9).value != expected_gap:
            checks.fail("R004", f"{section} 间隔公式缓存错误")
        expected_overlap = "N/A" if row == 2 else "NO"
        if sessions_v.cell(row, 10).value != expected_overlap:
            checks.fail("R004", f"{section} 重叠判定错误")
        if not formula_contains(sessions_f.cell(row, 5).value, "ROUND", "MOD"):
            checks.fail("R004", f"Session_Control!E{row} 不是时长公式")
        if not formula_contains(sessions_f.cell(row, 6).value, "COUNTIF", "ROSTER_CONTROL", f"A{row}"):
            checks.fail("R004", f"Session_Control!F{row} 不是人数公式")
        if row > 2 and not formula_contains(sessions_f.cell(row, 9).value, "ROUND", "1440"):
            checks.fail("R004", f"Session_Control!I{row} 不是间隔公式")
        session_context[section] = {
            "start": start,
            "end": end,
            "members": expected_names[section],
            "uid": section_uid(section),
        }

    reminders_f, reminders_v = formulas["Reminder_Control"], values["Reminder_Control"]
    expected_reminder_headers = [
        "Source Row", "Name", "Section", "Reserved Recipient", "Subject", "Message ID",
        "Calendar Attachment", "Session UID", "Delivery Status",
    ]
    if [reminders_f.cell(1, column).value for column in range(1, 10)] != expected_reminder_headers:
        checks.fail("R008", "Reminder_Control 表头不正确")
    for row, member in enumerate(members, start=2):
        expected_recipient = f"{member.name.lower()}@example.invalid"
        expected_subject = f"Team training reminder - {member.section} - 2020-05-08"
        expected_message_id = f"<{member.name.lower()}-training-20200508@evocowork.invalid>"
        expected_attachment = f"{section_slug(member.section)}-training.ics"
        observed_literals = [
            reminders_f.cell(row, 1).value,
            reminders_f.cell(row, 2).value,
            reminders_f.cell(row, 3).value,
            reminders_f.cell(row, 6).value,
            reminders_f.cell(row, 7).value,
            reminders_f.cell(row, 8).value,
            reminders_f.cell(row, 9).value,
        ]
        expected_literals = [
            member.source_row, member.name, member.section, expected_message_id, expected_attachment,
            section_uid(member.section), "BLOCKED_MISSING_REAL_EMAIL",
        ]
        if observed_literals != expected_literals:
            checks.fail("R008", f"Reminder_Control 第 {row} 行字段错误")
        if reminders_v.cell(row, 4).value != expected_recipient or reminders_v.cell(row, 5).value != expected_subject:
            checks.fail("R008", f"Reminder_Control 第 {row} 行收件人或主题错误")
        if not formula_contains(reminders_f.cell(row, 4).value, "LOWER", f"B{row}", "EXAMPLE.INVALID"):
            checks.fail("R004", f"Reminder_Control!D{row} 不是收件人公式")
        if not formula_contains(reminders_f.cell(row, 5).value, "TEAMTRAININGREMINDER", f"C{row}", "2020-05-08"):
            checks.fail("R004", f"Reminder_Control!E{row} 不是主题公式")

    summary_f, summary_v = formulas["Delivery_Summary"], values["Delivery_Summary"]
    expected_summary_headers = ["Metric", "Formula Value", "Expected", "Control Status", "Control Note"]
    if [summary_f.cell(1, column).value for column in range(1, 6)] != expected_summary_headers:
        checks.fail("R004", "Delivery_Summary 表头不正确")
    expected_metrics = [
        ("Source Rows", 12), ("Sections", 3), ("Reminder Messages", 12), ("Blocked Drafts", 12),
        ("Calendar Events", 3), ("Scheduled Hours", 6), ("Inter-session Gap (Minutes)", 60), ("Overlap Count", 0),
    ]
    for row, (metric, expected) in enumerate(expected_metrics, start=2):
        if summary_f.cell(row, 1).value != metric or summary_f.cell(row, 3).value != expected:
            checks.fail("R004", f"Delivery_Summary 第 {row} 行指标定义错误")
        if summary_v.cell(row, 2).value != expected or summary_v.cell(row, 4).value != "PASS":
            checks.fail("R004", f"Delivery_Summary 指标 {metric} 重算未通过")
        if not isinstance(summary_f.cell(row, 2).value, str) or not summary_f.cell(row, 2).value.startswith("="):
            checks.fail("R004", f"Delivery_Summary!B{row} 被静态化")

    formula_count = sum(
        1 for sheet in formulas.worksheets for row in sheet.iter_rows() for cell in row if cell.data_type == "f"
    )
    if formula_count < 88:
        checks.fail("R004", f"Excel 原生公式数不足：{formula_count}<88")
    expected_tables = {
        "Roster_Control": "TrainingRosterControlTable",
        "Session_Control": "TrainingSessionControlTable",
        "Reminder_Control": "TrainingReminderControlTable",
        "Delivery_Summary": "TrainingDeliverySummaryTable",
    }
    for sheet_name, table_name in expected_tables.items():
        sheet = formulas[sheet_name]
        if list(sheet.tables) != [table_name] or sheet.freeze_panes != "A2" or sheet.auto_filter.ref is None:
            checks.fail("R005", f"{sheet_name} 原生表格/筛选/冻结窗格不完整")
    if len(formulas["Delivery_Summary"]._charts) != 1:
        checks.fail("R005", "Delivery_Summary 必须有 1 个可编辑柱状图")
    conditional_ranges = sum(len(sheet.conditional_formatting) for sheet in formulas.worksheets)
    if conditional_ranges < 5:
        checks.fail("R005", f"条件格式范围不足：{conditional_ranges}<5")
    if any(len(sheet._images) for sheet in formulas.worksheets):
        checks.fail("R005", "Excel 不应用图片替代可编辑表格或图表")
    unsafe_formulas = [
        f"{sheet.title}!{cell.coordinate}"
        for sheet in formulas.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if cell.data_type == "f"
        and isinstance(cell.value, str)
        and ("http://" in cell.value.lower() or "https://" in cell.value.lower() or "[" in cell.value)
    ]
    if unsafe_formulas:
        checks.fail("R005", f"Excel 公式含外部引用：{unsafe_formulas}")
        checks.fail("R011", "Excel 含无来源外部引用")

    context["sessions"] = session_context
    context["names"] = [member.name for member in members]
    context["blocked"] = 12
    return context


def unfold_ics(raw: bytes) -> list[str]:
    if not raw.endswith(b"\r\n") or b"\n" in raw.replace(b"\r\n", b""):
        raise ValueError("ICS 必须使用 CRLF")
    text = raw.decode("utf-8")
    unfolded = re.sub(r"\r\n[ \t]", "", text)
    return [line for line in unfolded.split("\r\n") if line]


def unescape_ics(value: str) -> str:
    return value.replace("\\n", "\n").replace("\\,", ",").replace("\\;", ";").replace("\\\\", "\\")


def parse_calendar_bytes(raw: bytes) -> dict[str, Any]:
    lines = unfold_ics(raw)
    if not lines or lines[0] != "BEGIN:VCALENDAR" or lines[-1] != "END:VCALENDAR":
        raise ValueError("VCALENDAR 边界错误")
    if "VERSION:2.0" not in lines or "METHOD:PUBLISH" not in lines:
        raise ValueError("VCALENDAR VERSION/METHOD 错误")
    events: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    for line in lines:
        if line == "BEGIN:VEVENT":
            if current is not None:
                raise ValueError("VEVENT 嵌套")
            current = {}
            continue
        if line == "END:VEVENT":
            if current is None:
                raise ValueError("VEVENT 结束缺少开始")
            events.append(current)
            current = None
            continue
        if current is not None:
            if ":" not in line:
                raise ValueError(f"ICS 属性缺少冒号：{line}")
            raw_key, value = line.split(":", 1)
            key = raw_key.split(";", 1)[0].upper()
            if key in current:
                raise ValueError(f"VEVENT 属性重复：{key}")
            current[key] = unescape_ics(value)
            if ";TZID=" in raw_key.upper() or ";VALUE=DATE" in raw_key.upper():
                current[f"_{key}_PARAMS"] = raw_key
    if current is not None:
        raise ValueError("VEVENT 未结束")
    return {"lines": lines, "events": events}


def check_event(event: dict[str, str], section: str, grouped: dict[str, list[SourceMember]], checks: Checks, criterion: str) -> None:
    start, end = SESSION_TIMES[section]
    roster = " | ".join(member.name for member in grouped[section])
    expected = {
        "UID": section_uid(section),
        "DTSTAMP": "20200501T100000Z",
        "DTSTART": f"20200508T{start.strftime('%H%M%S')}",
        "DTEND": f"20200508T{end.strftime('%H%M%S')}",
        "SUMMARY": f"Team training - {section}",
        "STATUS": "TENTATIVE",
        "X-EVOCOWORK-SECTION": section,
        "X-EVOCOWORK-ROSTER": roster,
        "X-EVOCOWORK-TIME-BASIS": "FLOATING_LOCAL_TIME",
        "X-EVOCOWORK-EVIDENCE-GAPS": "REAL_EMAIL_TIMEZONE_LOCATION",
    }
    for key, value in expected.items():
        if event.get(key) != value:
            checks.fail(criterion, f"{section} 事件 {key} 错误")
    description = event.get("DESCRIPTION", "")
    for required in (roster, SOURCE_NAME, "Time zone", "location", "not provided"):
        if required.lower() not in description.lower():
            checks.fail(criterion, f"{section} DESCRIPTION 缺少 {required}")
    for banned in ("LOCATION", "ORGANIZER", "ATTENDEE", "TZID", "GEO"):
        if banned in event or f"_{banned}_PARAMS" in event:
            checks.fail("R011", f"{section} 伪造了 {banned}")
            checks.fail(criterion, f"{section} 含不允许属性 {banned}")
    if any("TZID=" in value.upper() for key, value in event.items() if key.startswith("_")):
        checks.fail("R011", f"{section} 伪造了 TZID 时区参数")
        checks.fail(criterion, f"{section} 必须使用 floating local time，不得绑定 TZID")


def check_master_calendar(path: Path, members: list[SourceMember], checks: Checks) -> dict[str, dict[str, str]]:
    grouped = grouped_members(members)
    event_map: dict[str, dict[str, str]] = {}
    try:
        parsed = parse_calendar_bytes(path.read_bytes())
        events = parsed["events"]
        if len(events) != 3:
            checks.fail("R006", f"Master ICS 应有 3 个 VEVENT，实际 {len(events)}")
        for event in events:
            section = event.get("X-EVOCOWORK-SECTION", "")
            if section not in SESSION_TIMES or section in event_map:
                checks.fail("R006", f"Master ICS Section 缺失或重复：{section}")
                continue
            check_event(event, section, grouped, checks, "R006")
            event_map[section] = event
        if set(event_map) != set(SESSION_TIMES):
            checks.fail("R006", "Master ICS 没有恰好覆盖 A/B/C")
        if any(line.startswith(("BEGIN:VALARM", "BEGIN:VTIMEZONE")) for line in parsed["lines"]):
            checks.fail("R011", "Master ICS 不得增加无来源时区或提醒")
    except Exception as exc:
        checks.fail("R006", f"Master ICS 无法解析：{exc}")
        checks.fail("R012", "Master ICS 无法参与跨文件对齐")
    return event_map


def message_factory(handle):
    return message_from_binary_file(handle, policy=policy.default)


def check_mbox(
    path: Path,
    members: list[SourceMember],
    master_events: dict[str, dict[str, str]],
    checks: Checks,
) -> dict[str, Any]:
    context: dict[str, Any] = {"messages": 0, "names": [], "uids": []}
    grouped = grouped_members(members)
    try:
        box = mailbox.mbox(path, factory=message_factory, create=False)
        messages = list(box)
    except Exception as exc:
        for criterion in ("R007", "R008", "R009", "R012", "R014"):
            checks.fail(criterion, f"MBOX 无法解析：{exc}")
        return context
    context["messages"] = len(messages)
    if len(messages) != 12:
        checks.fail("R007", f"MBOX 应有 12 封邮件，实际 {len(messages)}")
    by_name: dict[str, Any] = {}
    for message in messages:
        display_name, address = parseaddr(str(message.get("To", "")))
        if display_name in by_name:
            checks.fail("R007", f"MBOX 收件人重复：{display_name}")
        by_name[display_name] = message
        context["names"].append(display_name)
        if not message.is_multipart():
            checks.fail("R007", f"{display_name} 邮件不是 multipart MIME")
        if message.get("Cc") or message.get("Bcc"):
            checks.fail("R011", f"{display_name} 邮件增加了无来源 Cc/Bcc")
        if not address.endswith("@example.invalid"):
            checks.fail("R011", f"{display_name} 使用了可投递或非指定占位地址：{address}")
    for member in members:
        message = by_name.get(member.name)
        if message is None:
            checks.fail("R007", f"缺少 {member.name} 的独立提醒")
            checks.fail("R012", f"MBOX 未覆盖 {member.name}")
            continue
        expected_address = f"{member.name.lower()}@example.invalid"
        _, actual_address = parseaddr(str(message.get("To", "")))
        expected_subject = f"Team training reminder - {member.section} - 2020-05-08"
        expected_message_id = f"<{member.name.lower()}-training-20200508@evocowork.invalid>"
        expected_uid = section_uid(member.section)
        if actual_address != expected_address or str(message.get("Subject", "")) != expected_subject:
            checks.fail("R008", f"{member.name} 收件地址或主题错误")
        if str(message.get("Message-ID", "")) != expected_message_id:
            checks.fail("R008", f"{member.name} Message-ID 错误")
        if str(message.get("From", "")) != "Training Coordinator <training-coordinator@example.invalid>":
            checks.fail("R011", f"{member.name} 邮件 From 不是保留地址")
        if message.get("X-EvoCowork-Delivery-Status") != "BLOCKED_MISSING_REAL_EMAIL":
            checks.fail("R008", f"{member.name} 邮件未标记阻塞")
        if message.get("X-EvoCowork-Source-Row") != str(member.source_row):
            checks.fail("R014", f"{member.name} 邮件无法追回源行")
        if message.get("X-EvoCowork-Session-UID") != expected_uid:
            checks.fail("R008", f"{member.name} 邮件 Session UID 错误")
        body_part = message.get_body(preferencelist=("plain",))
        body = body_part.get_content() if body_part is not None else ""
        start, end = SESSION_TIMES[member.section]
        required_body = [
            f"Hello {member.name}", "Training date: 2020-05-08", f"Section: {member.section}",
            f"Time: {start.strftime('%H:%M')}-{end.strftime('%H:%M')}", "floating local time",
            "time zone not provided", "Location: not provided", expected_uid,
            "BLOCKED_MISSING_REAL_EMAIL", ".invalid", "draft", "has not been sent",
        ]
        for required in required_body:
            if required.lower() not in body.lower():
                checks.fail("R008", f"{member.name} 正文缺少 {required}")
        if re.search(r"\b(has|was|were) sent\b", body, flags=re.IGNORECASE) and "has not been sent" not in body.lower():
            checks.fail("R011", f"{member.name} 邮件声称已发送")
        attachments = list(message.iter_attachments())
        if len(attachments) != 1:
            checks.fail("R009", f"{member.name} 应恰好有 1 个日历附件，实际 {len(attachments)}")
            continue
        attachment = attachments[0]
        expected_filename = f"{section_slug(member.section)}-training.ics"
        if attachment.get_content_type() != "text/calendar" or attachment.get_filename() != expected_filename:
            checks.fail("R009", f"{member.name} 日历附件类型或文件名错误")
        raw_attachment = attachment.get_payload(decode=True) or b""
        try:
            parsed = parse_calendar_bytes(raw_attachment)
            if len(parsed["events"]) != 1:
                checks.fail("R009", f"{member.name} 附件应只有 1 个 VEVENT")
                continue
            event = parsed["events"][0]
            check_event(event, member.section, grouped, checks, "R009")
            master = master_events.get(member.section)
            compare_keys = (
                "UID", "DTSTART", "DTEND", "SUMMARY", "STATUS", "DESCRIPTION",
                "X-EVOCOWORK-SECTION", "X-EVOCOWORK-ROSTER", "X-EVOCOWORK-TIME-BASIS",
                "X-EVOCOWORK-EVIDENCE-GAPS",
            )
            if master is None or any(event.get(key) != master.get(key) for key in compare_keys):
                checks.fail("R009", f"{member.name} 附件与 Master ICS 不一致")
            context["uids"].append(event.get("UID"))
        except Exception as exc:
            checks.fail("R009", f"{member.name} 日历附件无法解析：{exc}")
    if set(by_name) != {member.name for member in members}:
        checks.fail("R007", "MBOX 收件人集与源表不一致")
    return context


def document_text(document: Document) -> str:
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def check_document(path: Path, members: list[SourceMember], checks: Checks) -> dict[str, Any]:
    context: dict[str, Any] = {}
    try:
        inspect_ooxml(path, checks)
        document = Document(path)
    except Exception as exc:
        for criterion in ("R010", "R011", "R012"):
            checks.fail(criterion, f"Word 无法打开：{exc}")
        return context
    text = document_text(document)
    required = [
        "Team Training Coordination Brief", "2020-05-08", "12 members", "3 sections",
        "12 reminder drafts blocked", "Section A", "09:00-11:00", "Section B", "11:00-13:00",
        "Section C", "14:00-16:00", "60-minute inter-session gap", "exactly two hours",
        "BLOCKED_MISSING_REAL_EMAIL", ".invalid", "floating local time", "No TZID",
        "No LOCATION", "No ORGANIZER or ATTENDEE", "not sent", "remain tentative",
    ]
    for item in required:
        if item.lower() not in text.lower():
            checks.fail("R010", f"Word 缺少关键信息：{item}")
    for member in members:
        if member.name not in text or member.source_time not in text or f"{member.name.lower()}@example.invalid" not in text:
            checks.fail("R010", f"Word 名单缺少或改写了 {member.name} 的源信息")
    if len(document.tables) != 4:
        checks.fail("R010", f"Word 应包含 4 个原生表格，实际 {len(document.tables)}")
    if document.inline_shapes:
        checks.fail("R010", "Word 不得用图片代替原生表格")
    page_width_cm = document.sections[0].page_width.cm
    page_height_cm = document.sections[0].page_height.cm
    if abs(page_width_cm - 21.0) > 0.2 or abs(page_height_cm - 29.7) > 0.2:
        checks.fail("R013", "Word 不是 A4 纵向页面")
    banned_patterns = [
        r"@[A-Za-z0-9.-]+\.(?:com|edu|org|net)\b",
        r"\b(?:GMT|PST|EST|CST)\b",
        r"\bRoom\s+\w+\b",
        r"\b(?:has|have|was|were) been sent\b",
    ]
    safe_text = re.sub(r"[A-Za-z0-9._%+-]+@(?:example|evocowork)\.invalid", "", text, flags=re.IGNORECASE)
    for pattern in banned_patterns:
        if re.search(pattern, safe_text, flags=re.IGNORECASE):
            checks.fail("R011", f"Word 包含无来源声明：{pattern}")
    context["text"] = text
    return context


def render_office(path: Path) -> int:
    with tempfile.TemporaryDirectory(prefix="training-judge-render-") as temporary:
        temporary_path = Path(temporary)
        output_dir = temporary_path / "output"
        profile_dir = temporary_path / "profile"
        output_dir.mkdir()
        process = subprocess.run(
            [
                "libreoffice", "--headless", f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "pdf", "--outdir", str(output_dir), str(path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=300,
        )
        pdf_path = output_dir / f"{path.stem}.pdf"
        if process.returncode != 0 or not pdf_path.is_file() or pdf_path.stat().st_size < 2000:
            raise RuntimeError(process.stdout[-1000:])
        info = subprocess.run(
            ["pdfinfo", str(pdf_path)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60
        )
        if info.returncode != 0:
            raise RuntimeError(info.stdout[-1000:])
        match = re.search(r"^Pages:\s+(\d+)\s*$", info.stdout, flags=re.MULTILINE)
        if not match:
            raise RuntimeError("pdfinfo 未返回页数")
        return int(match.group(1))


def check_cross_artifact(
    members: list[SourceMember],
    workbook_context: dict[str, Any],
    master_events: dict[str, dict[str, str]],
    mbox_context: dict[str, Any],
    document_context: dict[str, Any],
    checks: Checks,
) -> None:
    names = [member.name for member in members]
    if workbook_context.get("names") != names:
        checks.fail("R012", "Excel 人员顺序与源表不一致")
    if set(master_events) != set(SESSION_TIMES):
        checks.fail("R012", "Master ICS 三组覆盖不完整")
    if mbox_context.get("messages") != 12 or mbox_context.get("names") != names:
        checks.fail("R012", "MBOX 数量或顺序与 Excel/源表不一致")
    expected_uids = [section_uid(member.section) for member in members]
    if mbox_context.get("uids") != expected_uids:
        checks.fail("R012", "MBOX 附件 UID 顺序与人员 Section 不一致")
    word_text = document_context.get("text", "")
    if word_text:
        for member in members:
            if member.name not in word_text or member.section not in word_text:
                checks.fail("R012", f"Word 与源表对 {member.name} 的记录不一致")


def evaluate(artifact_dir: Path) -> dict[str, Any]:
    checks = Checks()
    task_root = Path(__file__).resolve().parents[1]
    try:
        members = read_source_truth(task_root)
    except Exception as exc:
        members = []
        checks.fail("R001", f"冻结输入真值无法读取：{exc}")

    if not artifact_dir.is_dir():
        checks.fail("R001", "交付目录不存在")
        actual_files: set[str] = set()
    else:
        actual_files = {path.name for path in artifact_dir.iterdir() if path.is_file()}
        extra_entries = [path.name for path in artifact_dir.iterdir() if not path.is_file()]
        if extra_entries:
            checks.fail("R001", f"交付目录含非文件项：{extra_entries}")
    missing = sorted(set(REQUIRED_FILES) - actual_files)
    extra = sorted(actual_files - set(REQUIRED_FILES))
    if missing or extra:
        checks.fail("R001", f"文件集不一致；缺失={missing}，额外={extra}")

    for name in REQUIRED_FILES:
        path = artifact_dir / name
        if name in actual_files and path.stat().st_size == 0:
            checks.fail("R001", f"{name} 是空文件")

    workbook_context: dict[str, Any] = {}
    master_events: dict[str, dict[str, str]] = {}
    mbox_context: dict[str, Any] = {}
    document_context: dict[str, Any] = {}
    if members and XLSX_NAME in actual_files:
        workbook_context = check_workbook(artifact_dir / XLSX_NAME, members, checks)
    else:
        for criterion in ("R002", "R003", "R004", "R005", "R012", "R014"):
            checks.fail(criterion, "Excel 交付不可评")
    if members and ICS_NAME in actual_files:
        master_events = check_master_calendar(artifact_dir / ICS_NAME, members, checks)
    else:
        checks.fail("R006", "Master ICS 交付不可评")
        checks.fail("R012", "Master ICS 缺失")
    if members and MBOX_NAME in actual_files:
        mbox_context = check_mbox(artifact_dir / MBOX_NAME, members, master_events, checks)
    else:
        for criterion in ("R007", "R008", "R009", "R012", "R014"):
            checks.fail(criterion, "MBOX 交付不可评")
    if members and DOCX_NAME in actual_files:
        document_context = check_document(artifact_dir / DOCX_NAME, members, checks)
    else:
        for criterion in ("R010", "R011", "R012"):
            checks.fail(criterion, "Word 交付不可评")
    if members:
        check_cross_artifact(members, workbook_context, master_events, mbox_context, document_context, checks)

    render_pages: dict[str, int] = {}
    core_ready = all(checks.passed(criterion) for criterion in [f"R{index:03d}" for index in range(1, 13)])
    if core_ready:
        for name, expected_pages in ((XLSX_NAME, 4), (DOCX_NAME, 2)):
            try:
                pages = render_office(artifact_dir / name)
                render_pages[name] = pages
                if pages != expected_pages:
                    checks.fail("R013", f"{name} 渲染页数应为 {expected_pages}，实际 {pages}")
            except Exception as exc:
                checks.fail("R013", f"{name} LibreOffice 渲染失败：{exc}")
    else:
        checks.fail("R013", "核心交付未通过，不授予视觉就绪分")

    hurdle_failures = [criterion for criterion in HURDLE_CAPS if not checks.passed(criterion)]
    base_score = sum(weight for criterion, weight in WEIGHTS.items() if checks.passed(criterion)) / 100.0
    cap = min([1.0] + [HURDLE_CAPS[criterion] for criterion in hurdle_failures])
    penalties: list[dict[str, Any]] = []
    penalty_total = 0.0
    if any(not checks.passed(criterion) for criterion in ("R002", "R003", "R004")):
        penalties.append({"id": "R015", "amount": 0.30, "reason": "人员/时间/公式真值错误"})
        penalty_total += 0.30
    if not checks.passed("R011"):
        penalties.append({"id": "R016", "amount": 0.25, "reason": "证据边界或安全状态错误"})
        penalty_total += 0.25
    if any(not checks.passed(criterion) for criterion in ("R001", "R005", "R007")):
        penalties.append({"id": "R017", "amount": 0.25, "reason": "文件/原生结构/MBOX 完整性错误"})
        penalty_total += 0.25
    if any(not checks.passed(criterion) for criterion in ("R006", "R008", "R009", "R010", "R012")):
        penalties.append({"id": "R018", "amount": 0.20, "reason": "跨产物一致性错误"})
        penalty_total += 0.20
    reward = max(0.0, min(base_score, cap) - penalty_total)
    reward = round(reward, 6)
    passed = reward >= 0.8 and not hurdle_failures
    return {
        "task_id": "cowork-evo/officebench-3-6-team-training-coordination-package",
        "reward": reward,
        "passed": passed,
        "base_score": round(base_score, 6),
        "hurdle_cap": cap,
        "hurdle_failures": hurdle_failures,
        "penalties": penalties,
        "criteria": checks.values,
        "source_truth": {
            "sha256": SOURCE_SHA256,
            "rows": len(members),
            "sections": {section: [member.name for member in grouped_members(members).get(section, [])] for section in SESSION_TIMES},
            "training_date": TRAINING_DATE.isoformat(),
        },
        "render_pages": render_pages,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = evaluate(args.artifact_dir)
    except Exception as exc:
        result = {
            "task_id": "cowork-evo/officebench-3-6-team-training-coordination-package",
            "reward": 0.0,
            "passed": False,
            "fatal_error": f"{type(exc).__name__}: {exc}",
            "hurdle_failures": ["R001"],
        }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{float(result['reward']):.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
