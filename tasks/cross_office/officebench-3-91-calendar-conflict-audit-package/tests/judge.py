#!/usr/bin/env python3
"""OfficeBench 3-91 日历冲突审计任务的独立确定性 Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-3-91-calendar-conflict-audit-package"
RUBRIC_VERSION = "officebench-3-91-calendar-conflict-audit-v0.1"
JUDGE_VERSION = "officebench-3-91-calendar-conflict-audit-deterministic-v0.1"
XLSX_NAME = "Calendar_Conflict_Audit.xlsx"
ICS_NAME = "Merged_Schedule.ics"
DOCX_NAME = "Calendar_Conflict_Report.docx"
REQUIRED_FILES = [XLSX_NAME, ICS_NAME, DOCX_NAME]
SHEETS = ["Source_Events", "Pairwise_Conflicts", "Conflict_Windows", "Joint_Free_Windows", "Summary"]
SOURCE_HASHES = {
    "Bob.ics": "a641a03140fc0903fb3d3a320642c146edc5689ffae06712c84cb4b5ee60be98",
    "Tom.ics": "4f946c13f4ee53deb6d74b70f73da42cc9a1063fc1183b7daca4a44bf69eba99",
}
SEARCH_START = datetime(2024, 5, 1, 8, 0, 0, tzinfo=timezone.utc)
SEARCH_END = datetime(2024, 5, 1, 20, 0, 0, tzinfo=timezone.utc)

SOURCE_HEADERS = ["Event_ID", "Owner", "Summary", "DTSTART_UTC", "DTEND_UTC", "Duration_Minutes", "UID", "Source_File", "Internal_Conflict_Count", "Cross_Owner_Conflict_Count"]
PAIR_HEADERS = ["Conflict_ID", "Bob_Event_ID", "Bob_Summary", "Tom_Event_ID", "Tom_Summary", "Overlap_Start_UTC", "Overlap_End_UTC", "Overlap_Minutes", "Same_Title"]
WINDOW_HEADERS = ["Window_ID", "Start_UTC", "End_UTC", "Duration_Minutes", "Active_Bob_Events", "Active_Tom_Events", "Pairwise_Conflict_Count"]
FREE_HEADERS = ["Window_ID", "Search_Window_Date", "Start_UTC", "End_UTC", "Duration_Minutes", "Is_Longest", "Recommendation"]
SUMMARY_LABELS = [
    "Total_Events", "Bob_Events", "Tom_Events", "Internal_Conflict_Pairs", "Cross_Owner_Pairwise_Conflicts",
    "Merged_Conflict_Windows", "Total_Conflict_Window_Minutes", "Joint_Free_Windows", "Total_Joint_Free_Minutes",
    "Longest_Joint_Free_Minutes", "Same_Title_Same_Time_Pairs", "Source_Files", "Merged_ICS_Events",
]
SUSPICIOUS_MARKERS = ["solution/golden", "judge.py", "expectations.json", "judge-result", "/root/", "/home/", "c:\\"]

SPECS = {
    "R001": ("完整交付物与冻结输入", 5, "hurdle", 0.0),
    "R002": ("十个来源事件", 12, "hurdle", 0.3),
    "R003": ("时长与冲突计数公式", 8, "hurdle", 0.35),
    "R004": ("无单人内部冲突", 5, "hurdle", 0.45),
    "R005": ("五个跨人事件对", 12, "hurdle", 0.4),
    "R006": ("四个合并冲突窗口", 8, "hurdle", 0.45),
    "R007": ("四个联合空闲窗口", 10, "hurdle", 0.45),
    "R008": ("公式汇总", 8, "hurdle", 0.45),
    "R009": ("Excel 原生审计结构", 5, "positive", None),
    "R010": ("合并 ICS 事件完整性", 8, "hurdle", 0.4),
    "R011": ("Word 结构与表格", 5, "hurdle", 0.5),
    "R012": ("Word 冲突与空闲真值", 6, "hurdle", 0.5),
    "R013": ("三产物一致性", 5, "hurdle", 0.4),
    "R014": ("可渲染与专业排版", 3, "positive", None),
    "R015": ("无事件、时间、公式或汇总错误", -30, "penalty", None),
    "R016": ("无编造事实或跨文件冲突", -25, "penalty", None),
    "R017": ("无缺失、损坏、静态化或调试泄漏", -25, "penalty", None),
    "R018": ("合并 ICS 无丢失或改期", -20, "penalty", None),
}


@dataclass(frozen=True)
class Event:
    owner: str
    summary: str
    start: datetime
    end: datetime
    uid: str
    source_file: str


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def approx(value: object, expected: float, tolerance: float = 1e-6) -> bool:
    try:
        return abs(float(value) - float(expected)) <= tolerance
    except (TypeError, ValueError):
        return False


def parse_utc(value: str) -> datetime:
    return datetime.strptime(value, "%Y%m%dT%H%M%SZ").replace(tzinfo=timezone.utc)


def format_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def as_utc(value: object) -> datetime | None:
    if isinstance(value, datetime):
        return value.replace(tzinfo=timezone.utc) if value.tzinfo is None else value.astimezone(timezone.utc)
    if isinstance(value, str):
        raw = value.strip()
        for pattern in ("%Y-%m-%d %H:%M:%S UTC", "%Y-%m-%d %H:%M UTC", "%Y-%m-%dT%H:%M:%SZ"):
            try:
                return datetime.strptime(raw, pattern).replace(tzinfo=timezone.utc)
            except ValueError:
                pass
    return None


def unfold_ics(text: str) -> list[str]:
    lines: list[str] = []
    for raw in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if raw.startswith((" ", "\t")) and lines:
            lines[-1] += raw[1:]
        elif raw:
            lines.append(raw)
    return lines


def parse_ics_events(path: Path, owner: str | None = None, output: bool = False) -> tuple[list[Event], dict[str, object]]:
    metadata: dict[str, object] = {"errors": [], "calendar": {}, "order": [], "raw": ""}
    try:
        raw_bytes = path.read_bytes()
        text = raw_bytes.decode("utf-8")
    except Exception as exc:
        metadata["errors"].append(f"read:{type(exc).__name__}:{exc}")  # type: ignore[index]
        return [], metadata
    metadata["raw"] = text
    if output and (b"\r\n" not in raw_bytes or raw_bytes.replace(b"\r\n", b"").find(b"\n") >= 0):
        metadata["errors"].append("not_strict_crlf")  # type: ignore[index]
    events: list[Event] = []
    current: dict[str, str] | None = None
    calendar: dict[str, str] = {}
    for line in unfold_ics(text):
        if line == "BEGIN:VEVENT":
            current = {}
            continue
        if line == "END:VEVENT":
            if current is None:
                metadata["errors"].append("unexpected_event_end")  # type: ignore[index]
                continue
            required = {"DTSTART", "DTEND", "SUMMARY", "UID"}
            if not required.issubset(current):
                metadata["errors"].append(f"missing_event_fields:{sorted(required-set(current))}")  # type: ignore[index]
                current = None
                continue
            try:
                event_owner = current.get("X-EVOCOWORK-OWNER", owner or "")
                source = current.get("X-EVOCOWORK-SOURCE", f"{owner}.ics" if owner else "")
                event = Event(event_owner, current["SUMMARY"], parse_utc(current["DTSTART"]), parse_utc(current["DTEND"]), current["UID"], source)
                events.append(event)
                metadata["order"].append(event.uid)  # type: ignore[index]
            except Exception as exc:
                metadata["errors"].append(f"event_parse:{type(exc).__name__}:{exc}")  # type: ignore[index]
            current = None
            continue
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        base = key.split(";", 1)[0].upper()
        if current is not None:
            current[base] = value
            if output and (base in {"DTSTART", "DTEND"} and ("TZID" in key.upper() or not value.endswith("Z"))):
                metadata["errors"].append(f"non_utc:{base}:{value}")  # type: ignore[index]
            if output and base in {"ATTENDEE", "LOCATION", "PRIORITY", "ORGANIZER"}:
                metadata["errors"].append(f"unsupported_property:{base}")  # type: ignore[index]
        elif base in {"VERSION", "PRODID", "CALSCALE", "METHOD", "X-WR-CALNAME"}:
            calendar[base] = value
    metadata["calendar"] = calendar
    return events, metadata


def overlap(left: Event, right: Event) -> tuple[datetime, datetime] | None:
    start = max(left.start, right.start)
    end = min(left.end, right.end)
    return (start, end) if start < end else None


def source_path(task_root: Path, filename: str) -> Path | None:
    """解析宿主机、Harbor 容器和兼容目录中的冻结输入。"""

    candidates = [
        Path("/app/input_files") / filename,
        task_root / "data" / "input_files" / filename,
        task_root / "environment" / "input_files" / filename,
    ]
    return next((path for path in candidates if path.is_file()), None)


def truth_from_sources(task_root: Path) -> tuple[dict[str, object], list[str]]:
    errors: list[str] = []
    events: list[Event] = []
    for filename, owner in (("Bob.ics", "Bob"), ("Tom.ics", "Tom")):
        path = source_path(task_root, filename)
        if path is None:
            errors.append(f"missing_source:{filename}")
            continue
        actual = sha256(path)
        if actual != SOURCE_HASHES[filename]:
            errors.append(f"source_hash:{filename}:{actual}")
            continue
        parsed, meta = parse_ics_events(path, owner=owner)
        errors.extend(f"{filename}:{item}" for item in meta["errors"])
        events.extend(parsed)
    events.sort(key=lambda event: (event.start, event.owner, event.summary, event.uid))
    ids = {event: f"E{index:02d}" for index, event in enumerate(events, start=1)}
    internal = [(left, right) for owner in ("Bob", "Tom") for index, left in enumerate([item for item in events if item.owner == owner]) for right in [item for item in events if item.owner == owner][index + 1 :] if overlap(left, right)]
    pairs = []
    for bob in [item for item in events if item.owner == "Bob"]:
        for tom in [item for item in events if item.owner == "Tom"]:
            interval = overlap(bob, tom)
            if interval:
                pairs.append((bob, tom, interval[0], interval[1]))
    pairs.sort(key=lambda row: (row[2], row[0].summary, row[1].summary))
    conflict_windows: list[tuple[datetime, datetime]] = []
    for _, _, start, end in pairs:
        if not conflict_windows or start > conflict_windows[-1][1]:
            conflict_windows.append((start, end))
        else:
            conflict_windows[-1] = (conflict_windows[-1][0], max(conflict_windows[-1][1], end))
    busy: list[tuple[datetime, datetime]] = []
    for event in events:
        start, end = max(event.start, SEARCH_START), min(event.end, SEARCH_END)
        if start >= end:
            continue
        if not busy or start > busy[-1][1]:
            busy.append((start, end))
        else:
            busy[-1] = (busy[-1][0], max(busy[-1][1], end))
    free: list[tuple[datetime, datetime]] = []
    cursor = SEARCH_START
    for start, end in busy:
        if cursor < start:
            free.append((cursor, start))
        cursor = max(cursor, end)
    if cursor < SEARCH_END:
        free.append((cursor, SEARCH_END))
    return {"events": events, "ids": ids, "internal": internal, "pairs": pairs, "conflict_windows": conflict_windows, "free": free}, errors


def inspect_package(path: Path) -> dict[str, object]:
    result: dict[str, object] = {"broken": False, "macros": [], "external": [], "missing_targets": [], "media": []}
    try:
        with zipfile.ZipFile(path) as package:
            names = set(package.namelist())
            result["media"] = sorted(name for name in names if "/media/" in name)
            result["macros"] = sorted(name for name in names if name.lower().endswith(("vbaproject.bin", ".exe", ".js")))
            for name in names:
                if not name.endswith(".rels"):
                    continue
                root = etree.fromstring(package.read(name))
                for relationship in root:
                    target = relationship.get("Target", "")
                    if relationship.get("TargetMode") == "External":
                        result["external"].append(f"{name}:{target}")
            bad = package.testzip()
            if bad:
                result["broken"] = True
                result["bad_member"] = bad
    except Exception as exc:
        result["broken"] = True
        result["error"] = f"{type(exc).__name__}:{exc}"
    return result


def recalculate_candidate(path: Path) -> tuple[Path | None, str]:
    root = Path(tempfile.mkdtemp(prefix="calendar-judge-recalc-"))
    source_dir, output_dir, profile = root / "source", root / "output", root / "profile"
    source_dir.mkdir(); output_dir.mkdir(); profile.mkdir()
    source = source_dir / path.name
    shutil.copy2(path, source)
    try:
        proc = subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "xlsx", "--outdir", str(output_dir), str(source)],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
        )
    except Exception as exc:
        return None, f"{type(exc).__name__}:{exc}"
    converted = output_dir / path.name
    return (converted, proc.stdout[-1000:]) if proc.returncode == 0 and converted.is_file() else (None, proc.stdout[-1000:])


def headers(sheet, count: int) -> list[str]:
    return [normalize(sheet.cell(1, column).value) for column in range(1, count + 1)]


def compare_event(got: list[object], expected: list[object]) -> bool:
    for index, (left, right) in enumerate(zip(got, expected, strict=True)):
        if index in {3, 4}:
            parsed = as_utc(left)
            if parsed is None or abs((parsed - right).total_seconds()) > 0.5:
                return False
        elif index in {5, 8, 9}:
            if not approx(left, float(right)):
                return False
        elif normalize(left) != normalize(right):
            return False
    return True


def parse_workbook(path: Path, truth: dict[str, object]) -> dict[str, object]:
    result: dict[str, object] = {
        "readable": False, "source_errors": [], "formula_errors": [], "internal_errors": [], "pair_errors": [],
        "window_errors": [], "free_errors": [], "summary_errors": [], "native": {},
    }
    try:
        formulas = load_workbook(path, data_only=False)
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}:{exc}"
        return result
    recalculated, log = recalculate_candidate(path)
    result["recalc_log"] = log
    if recalculated is None:
        return result
    try:
        values = load_workbook(recalculated, data_only=True)
    except Exception as exc:
        result["error"] = f"recalc_load:{type(exc).__name__}:{exc}"
        return result
    result["readable"] = True
    if formulas.sheetnames != SHEETS:
        result["source_errors"].append(f"sheetnames={formulas.sheetnames}")
        return result
    events: list[Event] = truth["events"]  # type: ignore[assignment]
    ids: dict[Event, str] = truth["ids"]  # type: ignore[assignment]
    pairs: list[tuple[Event, Event, datetime, datetime]] = truth["pairs"]  # type: ignore[assignment]
    conflict_windows: list[tuple[datetime, datetime]] = truth["conflict_windows"]  # type: ignore[assignment]
    free_windows: list[tuple[datetime, datetime]] = truth["free"]  # type: ignore[assignment]

    source_f, source_v = formulas["Source_Events"], values["Source_Events"]
    if headers(source_f, 10) != SOURCE_HEADERS or source_f.max_row != 11 or source_f.max_column != 10:
        result["source_errors"].append("source_structure")
    for row_index, event in enumerate(events, start=2):
        internal_count = sum(1 for other in events if other.owner == event.owner and other != event and overlap(event, other))
        cross_count = sum(1 for other in events if other.owner != event.owner and overlap(event, other))
        wanted = [ids[event], event.owner, event.summary, event.start, event.end, (event.end-event.start).total_seconds()/60, event.uid, event.source_file, internal_count, cross_count]
        got = [source_v.cell(row_index, column).value for column in range(1, 11)]
        if not compare_event(got, wanted):
            result["source_errors"].append(f"row{row_index}:{got}")
        for column in (6, 9, 10):
            if source_f.cell(row_index, column).data_type != "f":
                result["formula_errors"].append(f"Source_Events!{source_f.cell(row_index,column).coordinate}:not_formula")
        if internal_count != 0:
            result["internal_errors"].append(f"source_internal_truth:{ids[event]}")

    pair_f, pair_v = formulas["Pairwise_Conflicts"], values["Pairwise_Conflicts"]
    if headers(pair_f, 9) != PAIR_HEADERS or pair_f.max_row != 6 or pair_f.max_column != 9:
        result["pair_errors"].append("pair_structure")
    for row_index, (bob, tom, start, end) in enumerate(pairs, start=2):
        wanted = [f"P{row_index-1:02d}", ids[bob], bob.summary, ids[tom], tom.summary, start, end, (end-start).total_seconds()/60, "Yes" if bob.summary == tom.summary else "No"]
        got = [pair_v.cell(row_index, column).value for column in range(1, 10)]
        valid = True
        for index, (left, right) in enumerate(zip(got, wanted, strict=True)):
            if index in {5, 6}:
                parsed = as_utc(left); valid &= parsed is not None and abs((parsed-right).total_seconds()) <= 0.5
            elif index == 7:
                valid &= approx(left, right)
            else:
                valid &= normalize(left) == normalize(right)
        if not valid:
            result["pair_errors"].append(f"row{row_index}:{got}")
        for column in (8, 9):
            if pair_f.cell(row_index, column).data_type != "f":
                result["formula_errors"].append(f"Pairwise_Conflicts!{pair_f.cell(row_index,column).coordinate}:not_formula")

    window_f, window_v = formulas["Conflict_Windows"], values["Conflict_Windows"]
    if headers(window_f, 7) != WINDOW_HEADERS or window_f.max_row != 5 or window_f.max_column != 7:
        result["window_errors"].append("window_structure")
    for row_index, (start, end) in enumerate(conflict_windows, start=2):
        bob_names = "; ".join(sorted({event.summary for event in events if event.owner == "Bob" and event.start < end and event.end > start}))
        tom_names = "; ".join(sorted({event.summary for event in events if event.owner == "Tom" and event.start < end and event.end > start}))
        count = sum(1 for _, _, pair_start, pair_end in pairs if pair_start < end and pair_end > start)
        wanted = [f"W{row_index-1:02d}", start, end, (end-start).total_seconds()/60, bob_names, tom_names, count]
        got = [window_v.cell(row_index, column).value for column in range(1, 8)]
        valid = normalize(got[0]) == wanted[0] and normalize(got[4]) == wanted[4] and normalize(got[5]) == wanted[5] and approx(got[3], wanted[3]) and approx(got[6], wanted[6])
        valid &= as_utc(got[1]) is not None and abs((as_utc(got[1])-start).total_seconds()) <= 0.5  # type: ignore[operator]
        valid &= as_utc(got[2]) is not None and abs((as_utc(got[2])-end).total_seconds()) <= 0.5  # type: ignore[operator]
        if not valid:
            result["window_errors"].append(f"row{row_index}:{got}")
        for column in (4, 7):
            if window_f.cell(row_index, column).data_type != "f":
                result["formula_errors"].append(f"Conflict_Windows!{window_f.cell(row_index,column).coordinate}:not_formula")

    free_f, free_v = formulas["Joint_Free_Windows"], values["Joint_Free_Windows"]
    if headers(free_f, 7) != FREE_HEADERS or free_f.max_row != 5 or free_f.max_column != 7:
        result["free_errors"].append("free_structure")
    for row_index, (start, end) in enumerate(free_windows, start=2):
        minutes = (end-start).total_seconds()/60
        wanted = [f"F{row_index-1:02d}", "2024-05-01", start, end, minutes, "Yes" if minutes == 90 else "No", "Recommended longest joint window" if minutes == 90 else "Available alternative"]
        got = [free_v.cell(row_index, column).value for column in range(1, 8)]
        valid = normalize(got[0]) == wanted[0] and normalize(got[1]) == wanted[1] and approx(got[4], wanted[4]) and normalize(got[5]) == wanted[5] and normalize(got[6]) == wanted[6]
        valid &= as_utc(got[2]) is not None and abs((as_utc(got[2])-start).total_seconds()) <= 0.5  # type: ignore[operator]
        valid &= as_utc(got[3]) is not None and abs((as_utc(got[3])-end).total_seconds()) <= 0.5  # type: ignore[operator]
        if not valid:
            result["free_errors"].append(f"row{row_index}:{got}")
        for column in (5, 6, 7):
            if free_f.cell(row_index, column).data_type != "f":
                result["formula_errors"].append(f"Joint_Free_Windows!{free_f.cell(row_index,column).coordinate}:not_formula")

    summary_f, summary_v = formulas["Summary"], values["Summary"]
    labels = [normalize(summary_v.cell(row, 1).value) for row in range(2, 15)]
    expected_summary = [10, 5, 5, 0, 5, 4, 420, 4, 160, 90, 2, 2, 10]
    if labels != SUMMARY_LABELS:
        result["summary_errors"].append(f"labels={labels}")
    for row_index, expected in enumerate(expected_summary, start=2):
        if summary_f.cell(row_index, 2).data_type != "f":
            result["formula_errors"].append(f"Summary!B{row_index}:not_formula")
        if not approx(summary_v.cell(row_index, 2).value, expected):
            result["summary_errors"].append(f"B{row_index}={summary_v.cell(row_index,2).value!r}")
    for sheet in values.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and re.match(r"^#(?:NAME|VALUE|REF|DIV/0|N/A)", cell.value, flags=re.IGNORECASE):
                    result["formula_errors"].append(f"{sheet.title}!{cell.coordinate}={cell.value}")
    result["native"] = {
        "tables": sum(len(sheet.tables) for sheet in formulas.worksheets),
        "charts": sum(len(sheet._charts) for sheet in formulas.worksheets),
        "freeze": all(sheet.freeze_panes == "A2" for sheet in formulas.worksheets),
        "conditional_formats": sum(len(sheet.conditional_formatting) for sheet in formulas.worksheets),
    }
    return result


def parse_output_ics(path: Path, truth: dict[str, object]) -> dict[str, object]:
    result: dict[str, object] = {"errors": [], "event_errors": [], "order_errors": []}
    events, meta = parse_ics_events(path, output=True)
    result["errors"].extend(meta["errors"])
    calendar = meta["calendar"]
    for key, expected in (("VERSION", "2.0"), ("CALSCALE", "GREGORIAN"), ("METHOD", "PUBLISH"), ("X-WR-CALNAME", "Bob and Tom Merged Schedule")):
        if calendar.get(key) != expected:
            result["errors"].append(f"calendar:{key}={calendar.get(key)!r}")
    source_events: list[Event] = truth["events"]  # type: ignore[assignment]
    expected = {(event.uid, event.owner, event.source_file, event.summary, event.start, event.end) for event in source_events}
    got = {(event.uid, event.owner, event.source_file, event.summary, event.start, event.end) for event in events}
    if got != expected or len(events) != 10 or len({event.uid for event in events}) != 10:
        result["event_errors"].append({"expected": sorted((uid, owner, source, summary, format_utc(start), format_utc(end)) for uid, owner, source, summary, start, end in expected), "got": sorted((uid, owner, source, summary, format_utc(start), format_utc(end)) for uid, owner, source, summary, start, end in got), "count": len(events)})
    expected_order = [event.uid for event in source_events]
    if meta["order"] != expected_order:
        result["order_errors"].append({"expected": expected_order, "got": meta["order"]})
    result["events"] = len(events)
    return result


def word_text(document: Document) -> str:
    parts = [paragraph.text for paragraph in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            parts.extend(cell.text for cell in row.cells)
    return "\n".join(parts)


def unsupported_claims(text: str) -> list[str]:
    patterns = {
        "timezone_conversion": r"(?:converted?|translated?)\s+(?:to|into)\s+(?:EST|EDT|PST|PDT|CST|local time)",
        "invented_location": r"(?:location|room|venue)\s*(?:is|:|=)\s*(?:room|office|building|zoom)",
        "invented_priority": r"(?:priority|importance)\s*(?:is|:|=)\s*(?:high|low|urgent)",
        "created_meeting": r"(?:created|scheduled|added)\s+(?:a\s+)?(?:new|additional)?\s*(?:meeting|event|appointment)",
        "invented_attendee": r"(?:attendee|participant|organizer)\s*(?:is|:|=)",
    }
    return [name for name, pattern in patterns.items() if re.search(pattern, text, flags=re.IGNORECASE)]


def parse_word(path: Path, truth: dict[str, object]) -> dict[str, object]:
    result: dict[str, object] = {"readable": False, "structure_errors": [], "fact_errors": [], "unsupported": [], "tables": 0, "headings": []}
    try:
        document = Document(path)
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}:{exc}"
        return result
    result["readable"] = True
    result["tables"] = len(document.tables)
    result["headings"] = [normalize(paragraph.text) for paragraph in document.paragraphs if str(paragraph.style.name).startswith("Heading")]
    text = normalize(word_text(document))
    lower = text.lower()
    if "Calendar Conflict Audit Report" not in text:
        result["structure_errors"].append("wrong_title")
    for heading in ("Scope and Time Semantics", "Source Event Reconciliation", "Conflict Findings", "Joint Availability", "Merged Calendar Control", "Audit Conclusion"):
        if heading not in result["headings"]:
            result["structure_errors"].append(f"missing_heading:{heading}")
    structure_tokens = ["half-open intervals", "touching endpoints are not conflicts", "without local-time conversion", "X-EVOCOWORK-OWNER", "X-EVOCOWORK-SOURCE"]
    for token in structure_tokens:
        if token.lower() not in lower:
            result["structure_errors"].append(f"missing:{token}")
    fact_tokens = [
        "10 VEVENT", "five for Bob", "five for Tom", "0 internal conflict pairs", "five event pairs", "four merged conflict windows",
        "420 minutes", "four windows totaling 160 minutes", "14:30-16:00 UTC", "90 minutes", "class at 16:00-18:00 UTC",
        "dinner at 18:30-19:30 UTC", "No timezone, attendee, location, priority or additional meeting is inferred",
    ]
    for token in fact_tokens:
        if token.lower() not in lower:
            result["fact_errors"].append(f"missing:{token}")
    events: list[Event] = truth["events"]  # type: ignore[assignment]
    ids: dict[Event, str] = truth["ids"]  # type: ignore[assignment]
    for event in events:
        for token in (ids[event], event.owner, event.summary, event.start.strftime("%H:%M"), event.end.strftime("%H:%M"), event.source_file):
            if token.lower() not in lower:
                result["fact_errors"].append(f"missing_event:{ids[event]}:{token}")
    for token in ("P01", "P02", "P03", "P04", "P05", "W01", "W02", "W03", "W04", "F01", "F02", "F03", "F04"):
        if token.lower() not in lower:
            result["fact_errors"].append(f"missing_row:{token}")
    result["unsupported"] = unsupported_claims(text)
    return result


def render_office(artifact_dir: Path) -> dict[str, object]:
    result: dict[str, object] = {"xlsx_pages": 0, "docx_pages": 0, "errors": [], "logs": []}
    with tempfile.TemporaryDirectory(prefix="calendar-judge-render-") as tmp:
        root = Path(tmp); output = root / "output"; profile = root / "profile"; output.mkdir(); profile.mkdir()
        for name in (XLSX_NAME, DOCX_NAME):
            source = artifact_dir / name
            if not source.is_file():
                result["errors"].append(f"missing:{name}")
                continue
            try:
                proc = subprocess.run(
                    ["soffice", f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", str(output), str(source)],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300,
                )
            except Exception as exc:
                result["errors"].append(f"{name}:{type(exc).__name__}:{exc}")
                continue
            result["logs"].append(proc.stdout[-1000:])
            pdf = output / f"{Path(name).stem}.pdf"
            if proc.returncode != 0 or not pdf.is_file():
                result["errors"].append(f"{name}:conversion_failed")
                continue
            info = subprocess.run(["pdfinfo", str(pdf)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60).stdout
            match = re.search(r"^Pages:\s+(\d+)", info, flags=re.MULTILINE)
            pages = int(match.group(1)) if match else 0
            extracted = subprocess.run(["pdftotext", str(pdf), "-"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=60).stdout
            if name == XLSX_NAME:
                result["xlsx_pages"] = pages
                tokens = ("E01", "P05", "W04", "F04", "Total_Conflict_Window_Minutes", "Conflict vs joint free minutes")
            else:
                result["docx_pages"] = pages
                tokens = ("Calendar Conflict Audit Report", "Conflict Findings", "Joint Availability", "Merged Calendar Control", "Audit Conclusion")
            for token in tokens:
                if token not in extracted:
                    result["errors"].append(f"{name}:render_missing:{token}")
            if re.search(r"#(?:NAME|VALUE|REF|DIV/0|N/A)", extracted, flags=re.IGNORECASE):
                result["errors"].append(f"{name}:formula_error_visible")
    if result["xlsx_pages"] != 5:
        result["errors"].append(f"xlsx_pages={result['xlsx_pages']}")
    if result["docx_pages"] != 3:
        result["errors"].append(f"docx_pages={result['docx_pages']}")
    return result


def suspicious_content(artifact_dir: Path) -> list[str]:
    findings = []
    for path in artifact_dir.iterdir() if artifact_dir.is_dir() else []:
        if not path.is_file():
            continue
        try:
            if path.suffix.lower() in {".md", ".txt", ".ics"}:
                text = path.read_text(encoding="utf-8", errors="ignore").lower()
            elif path.suffix.lower() in {".xlsx", ".docx"}:
                with zipfile.ZipFile(path) as package:
                    text = " ".join(package.read(name).decode("utf-8", errors="ignore") for name in package.namelist() if name.endswith(".xml")).lower()
            else:
                text = ""
        except Exception:
            continue
        for marker in SUSPICIOUS_MARKERS:
            if marker in text:
                findings.append(f"{path.name}:{marker}")
    return sorted(set(findings))


def criterion(identifier: str, score: float, evidence: object) -> dict[str, object]:
    description, weight, kind, cap = SPECS[identifier]
    raw: dict[str, object] = {"description": description, "weight": weight, "type": kind, "passed": score >= 0.999}
    if cap is not None:
        raw["score_cap_on_failure"] = cap
    return {"criterion_id": identifier, "score": round(float(score), 6), "evidence": evidence, "method": "deterministic", "raw": raw}


def evaluate(artifact_dir: Path) -> dict[str, object]:
    task_root = Path(__file__).resolve().parents[1]
    truth, source_errors = truth_from_sources(task_root)
    delivered = sorted(path.name for path in artifact_dir.iterdir() if path.is_file()) if artifact_dir.is_dir() else []
    exact_files = delivered == sorted(REQUIRED_FILES)
    packages = {name: inspect_package(artifact_dir / name) for name in (XLSX_NAME, DOCX_NAME)}
    package_ok = all(not info["broken"] and not info["macros"] and not info["external"] for info in packages.values())
    workbook = parse_workbook(artifact_dir / XLSX_NAME, truth) if (artifact_dir / XLSX_NAME).is_file() and not source_errors else {
        "readable": False, "source_errors": ["unavailable"], "formula_errors": [], "internal_errors": [], "pair_errors": [], "window_errors": [], "free_errors": [], "summary_errors": [], "native": {},
    }
    calendar = parse_output_ics(artifact_dir / ICS_NAME, truth) if (artifact_dir / ICS_NAME).is_file() and not source_errors else {"errors": ["unavailable"], "event_errors": [], "order_errors": [], "events": 0}
    word = parse_word(artifact_dir / DOCX_NAME, truth) if (artifact_dir / DOCX_NAME).is_file() and not source_errors else {"readable": False, "structure_errors": ["unavailable"], "fact_errors": [], "unsupported": [], "tables": 0, "headings": []}
    render = render_office(artifact_dir) if exact_files else {"xlsx_pages": 0, "docx_pages": 0, "errors": ["incomplete"]}
    suspicious = suspicious_content(artifact_dir)

    source_ok = workbook.get("readable") and not workbook.get("source_errors")
    formulas_ok = source_ok and not workbook.get("formula_errors")
    internal_ok = formulas_ok and not workbook.get("internal_errors")
    pairs_ok = internal_ok and not workbook.get("pair_errors")
    windows_ok = pairs_ok and not workbook.get("window_errors")
    free_ok = windows_ok and not workbook.get("free_errors")
    summary_ok = free_ok and not workbook.get("summary_errors")
    native = workbook.get("native", {})
    native_ok = native.get("tables", 0) >= 5 and native.get("charts", 0) >= 1 and native.get("freeze") is True and native.get("conditional_formats", 0) >= 5
    calendar_ok = not calendar.get("errors") and not calendar.get("event_errors") and not calendar.get("order_errors") and calendar.get("events") == 10
    word_structure_ok = word.get("readable") and word.get("tables", 0) >= 4 and not word.get("structure_errors") and not packages[DOCX_NAME].get("media")
    word_facts_ok = word_structure_ok and not word.get("fact_errors")
    cross_ok = summary_ok and calendar_ok and word_facts_ok and not word.get("unsupported")
    safe_ok = exact_files and package_ok and not suspicious
    render_ok = not render.get("errors")

    results = [
        criterion("R001", 1.0 if exact_files and not source_errors and package_ok else 0.0, {"source_errors": source_errors, "delivered": delivered, "packages": packages}),
        criterion("R002", 1.0 if source_ok else 0.0, workbook.get("source_errors")),
        criterion("R003", 1.0 if formulas_ok else 0.0, workbook.get("formula_errors")),
        criterion("R004", 1.0 if internal_ok else 0.0, workbook.get("internal_errors")),
        criterion("R005", 1.0 if pairs_ok else 0.0, workbook.get("pair_errors")),
        criterion("R006", 1.0 if windows_ok else 0.0, workbook.get("window_errors")),
        criterion("R007", 1.0 if free_ok else 0.0, workbook.get("free_errors")),
        criterion("R008", 1.0 if summary_ok else 0.0, workbook.get("summary_errors")),
        criterion("R009", 1.0 if native_ok else 0.0, native),
        criterion("R010", 1.0 if calendar_ok else 0.0, calendar),
        criterion("R011", 1.0 if word_structure_ok else 0.0, {"tables": word.get("tables"), "headings": word.get("headings"), "errors": word.get("structure_errors")}),
        criterion("R012", 1.0 if word_facts_ok else 0.0, word.get("fact_errors")),
        criterion("R013", 1.0 if cross_ok else 0.0, {"workbook": summary_ok, "calendar": calendar_ok, "word": word_facts_ok, "unsupported": word.get("unsupported")}),
        criterion("R014", 1.0 if render_ok and native_ok else 0.0, {"render": render, "native": native}),
        criterion("R015", 1.0 if source_ok and formulas_ok and internal_ok and pairs_ok and windows_ok and free_ok and summary_ok else 0.0, {"source": source_ok, "formulas": formulas_ok, "internal": internal_ok, "pairs": pairs_ok, "windows": windows_ok, "free": free_ok, "summary": summary_ok}),
        criterion("R016", 1.0 if cross_ok and not word.get("unsupported") else 0.0, {"cross": cross_ok, "unsupported": word.get("unsupported")}),
        criterion("R017", 1.0 if safe_ok and formulas_ok and native_ok else 0.0, {"safe": safe_ok, "formulas": formulas_ok, "native": native_ok, "suspicious": suspicious}),
        criterion("R018", 1.0 if calendar_ok else 0.0, calendar),
    ]
    positive_points = sum(SPECS[item["criterion_id"]][1] * float(item["score"]) for item in results if SPECS[item["criterion_id"]][1] > 0)
    penalty_points = sum(SPECS[item["criterion_id"]][1] for item in results if SPECS[item["criterion_id"]][1] < 0 and float(item["score"]) < 0.999)
    reward = max(0.0, min(1.0, (positive_points + penalty_points) / 100.0))
    hurdle_failures = []
    for item in results:
        identifier = item["criterion_id"]
        _, _, kind, cap = SPECS[identifier]
        if kind == "hurdle" and float(item["score"]) < 0.999:
            hurdle_failures.append(identifier)
            reward = min(reward, float(cap))
    reward = round(reward, 6)
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [sha256(artifact_dir / name) for name in REQUIRED_FILES if (artifact_dir / name).is_file()],
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
