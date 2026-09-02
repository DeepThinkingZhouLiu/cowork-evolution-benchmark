#!/usr/bin/env python3
"""将 AutomationBench 1202 的三类公开状态静态化为本题输入。

本脚本只识别 ``sales.cross_modal_renewal_processing``，并对邮件、
续约表、Salesforce 联系人和 Slack 频道做逐字段断言。它不是通用
制题器，保留仅为了让这一道题的静态化过程可审计。
"""

from __future__ import annotations

import importlib.util
import json
import shutil
import sys
import types
import zipfile
from datetime import datetime
from email.generator import BytesGenerator
from email.message import EmailMessage
from email.policy import SMTP
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


ROOT = Path(__file__).resolve().parents[1]
UPSTREAM = Path(__file__).with_name("upstream-sales-tasks.py")
SNAPSHOT = Path(__file__).with_name("upstream-task-1202.json")
DATA_DIR = ROOT / "data" / "input_files"
ENV_DIR = ROOT / "environment" / "input_files"
FIXED_ZIP_TIME = (2026, 2, 24, 10, 15, 0)


def load_upstream_task() -> dict:
    datasets = types.ModuleType("datasets")
    datasets.Dataset = object
    sys.modules["datasets"] = datasets
    for name in ("automationbench", "automationbench.domains", "automationbench.domains.sales"):
        sys.modules[name] = types.ModuleType(name)
    noise = types.ModuleType("automationbench.domains.sales._noise")
    noise.apply_noise = lambda value: value
    sys.modules[noise.__name__] = noise

    spec = importlib.util.spec_from_file_location("automationbench_sales_tasks_frozen", UPSTREAM)
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入冻结的 AutomationBench 源文件")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    task = module.get_cross_modal_renewal_processing_task()
    if task.get("example_id") != 1202 or task.get("task") != "sales.cross_modal_renewal_processing":
        raise ValueError("冻结函数不再是 AutomationBench 1202")
    return task


def assert_task_shape(task: dict) -> None:
    state = task["info"]["initial_state"]
    assert [item["id"] for item in state["gmail"]["messages"]] == [
        "msg_cs_renewal_approval",
        "msg_external_partner",
    ]
    spreadsheet = state["google_sheets"]["spreadsheets"]
    assert len(spreadsheet) == 1 and spreadsheet[0]["id"] == "ss_renewals"
    rows = spreadsheet[0]["worksheets"][0]["rows"]
    assert [row["cells"]["Account ID"] for row in rows] == [
        "001xx000003REN1",
        "001xx000003REN2",
        "001xx000003REN3",
        "001xx000003REN4",
    ]
    accounts = state["salesforce"]["accounts"]
    contacts = state["salesforce"]["contacts"]
    assert len(accounts) == 7 and len(contacts) == 8
    assert sum(bool(contact.get("is_primary")) for contact in contacts) == 7
    channels = state["slack"]["channels"]
    assert len(channels) == 1 and channels[0] == {
        "id": "C_RENEWALS", "name": "renewals", "is_private": False
    }
    assert len(task["info"]["assertions"]) == 11


def message_bytes(record: dict) -> bytes:
    message = EmailMessage(policy=SMTP)
    message["From"] = record["from_"]
    message["To"] = ", ".join(record["to"])
    message["Subject"] = record["subject"]
    parsed = datetime.fromisoformat(record["date"].replace("Z", "+00:00"))
    message["Date"] = parsed.strftime("%a, %d %b %Y %H:%M:%S +0000")
    message["Message-ID"] = f"<{record['id']}@automationbench.invalid>"
    message["X-AutomationBench-Message-ID"] = record["id"]
    message["X-AutomationBench-Thread-ID"] = record["thread_id"]
    message["X-AutomationBench-Labels"] = ",".join(record.get("label_ids", []))
    message.set_content(record["body_plain"], charset="utf-8")
    buffer = BytesIO()
    BytesGenerator(buffer, policy=SMTP, mangle_from_=True, maxheaderlen=0).flatten(message)
    return buffer.getvalue()


def write_mbox(messages: list[dict], target: Path) -> None:
    chunks: list[bytes] = []
    for record in messages:
        parsed = datetime.fromisoformat(record["date"].replace("Z", "+00:00"))
        chunks.append(
            f"From {record['from_']} {parsed.strftime('%a %b %d %H:%M:%S %Y')}\n".encode("ascii")
        )
        chunks.append(message_bytes(record).replace(b"\r\n", b"\n"))
        chunks.append(b"\n")
    target.write_bytes(b"".join(chunks))


def normalize_xlsx(path: Path) -> None:
    temporary = path.with_suffix(".normalized.xlsx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(
        temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
    ) as destination:
        for name in sorted(source.namelist()):
            info = zipfile.ZipInfo(name, FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            destination.writestr(info, source.read(name))
    temporary.replace(path)


def write_tracker(state: dict, target: Path) -> None:
    source = state["google_sheets"]["spreadsheets"][0]
    worksheet = source["worksheets"][0]
    headers = ["Account", "Account ID", "Renewal Date", "Contract Value", "Notes"]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = worksheet["title"]
    workbook.creator = "AutomationBench 1202 public fixture"
    workbook.title = source["title"]
    workbook.subject = "Static Office representation of Google Sheets renewal state"
    workbook.description = "No accounts or values were added beyond AutomationBench 1202."
    sheet.append(headers)
    for item in worksheet["rows"]:
        cells = item["cells"]
        sheet.append([cells.get(header, "") for header in headers])

    table = Table(displayName="RenewalTrackerInput", ref=f"A1:E{sheet.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    sheet.add_table(table)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    widths = {"A": 28, "B": 24, "C": 16, "D": 18, "E": 44}
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="17365D")
        cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Aptos", size=10)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = sheet.dimensions
    workbook.save(target)
    normalize_xlsx(target)


def main() -> None:
    task = load_upstream_task()
    assert_task_shape(task)
    state = task["info"]["initial_state"]
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ENV_DIR.mkdir(parents=True, exist_ok=True)

    SNAPSHOT.write_text(json.dumps(task, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    write_mbox(state["gmail"]["messages"], DATA_DIR / "Renewal_Inbox.mbox")
    write_tracker(state, DATA_DIR / "Renewal_Tracker.xlsx")
    (DATA_DIR / "Salesforce_Renewal_Contacts.json").write_text(
        json.dumps(state["salesforce"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    (DATA_DIR / "Renewals_Channel.json").write_text(
        json.dumps(state["slack"], ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    for source in sorted(DATA_DIR.iterdir()):
        if source.is_file():
            shutil.copy2(source, ENV_DIR / source.name)


if __name__ == "__main__":
    main()
