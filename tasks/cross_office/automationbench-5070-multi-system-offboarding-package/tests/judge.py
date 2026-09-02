#!/usr/bin/env python3
"""AutomationBench 5070 本周范围、离职动作草稿与零副作用专属 Judge。"""

import argparse
import json
import mailbox
import re
from pathlib import Path

from openpyxl import load_workbook

from office_checks import body_text, docx_text, package_safe, pptx_text, render, sha256


TASK = "cowork-evo/automationbench-5070-multi-system-offboarding-package"
HASHES = {
    "Initial_Application_State.json": "ae2fb67332f1613150faaf09df20b906be87de8609382b7d4482fcc1e3a1e1dc",
    "Offboarding_Master.xlsx": "bf25180c05172e57d11eefc87a0f522ea18384f645b3426f857e641d548dd742",
    "Workflow_Context.json": "a11d79dcd287ae48fed36f167220293cf0a95200000012e2135aa9b9abe5e476",
}
XLSX = "Offboarding_Multi_System_Audit.xlsx"
DOCX = "Offboarding_Operations_Brief.docx"
PPTX = "Offboarding_Operations_Review.pptx"
BAMBOO = "BambooHR_Status_Update_Drafts.json"
JIRA = "Jira_Access_Revocation_Drafts.json"
PAYROLL = "Payroll_Notification_Draft.mbox"
TRACKER = "Tracker_Status_Update_Drafts.json"
CONTROL = "Offboarding_Execution_Control.json"
REQUIRED = {XLSX, DOCX, PPTX, BAMBOO, JIRA, PAYROLL, TRACKER, CONTROL}
WEIGHTS = {f"R{i:03d}": value for i, value in enumerate([6, 8, 10, 10, 9, 8, 9, 9, 8, 8, 7, 8], 1)}
CAPS = {key: (0 if key == "R001" else 0.25) for key in WEIGHTS if key != "R012"}


class Checks:
    def __init__(self):
        self.values = {key: {"passed": True, "details": []} for key in WEIGHTS}

    def require(self, key, condition, detail):
        if not condition:
            self.values[key]["passed"] = False
            if detail not in self.values[key]["details"]:
                self.values[key]["details"].append(detail)

    def ok(self, key):
        return self.values[key]["passed"]


def task_root():
    return Path(__file__).resolve().parents[1]


def input_dir():
    container = Path("/app/input_files")
    return container if container.is_dir() else task_root() / "data/input_files"


def rows(ws):
    return [tuple(cell.value for cell in row) for row in ws.iter_rows(min_row=2)]


def verify_source(checks):
    source = input_dir()
    for name, expected in HASHES.items():
        checks.require("R001", (source / name).is_file() and sha256(source / name) == expected, f"输入 {name} 缺失或哈希错误")

    wb = load_workbook(source / "Offboarding_Master.xlsx", read_only=True, data_only=True)
    observed = rows(wb["Week of March 18"])
    expected = [
        (2, "Greg Foster", "EMP-2001", "greg.foster@company.example.com", "2026-03-21", "Voluntary", "5 days", "Pending"),
        (3, "Ian Costa", "EMP-2004", "ian.costa@company.example.com", "2026-03-20", "Voluntary", "0 days", "Pending"),
        (4, "Nadia Petrova", "EMP-2005", "nadia.petrova@company.example.com", "2026-04-05", "Voluntary", "3 days", "Pending"),
    ]
    normalized = [(r[0], r[1], r[2], r[3], r[4].strftime("%Y-%m-%d"), r[5], r[6], r[7]) for r in observed]
    checks.require("R002", normalized == expected, "3 人源清单、日期、PTO 或状态不一致")
    convention = rows(wb["Offboarding Convention"])
    checks.require("R002", convention == [(2, "All required offboarding actions completed", "Processed")], "Processed 约定不一致")
    context = json.loads((source / "Workflow_Context.json").read_text())
    checks.require("R002", context.get("scope_start") == "2026-03-18" and context.get("scope_end") == "2026-03-22" and context.get("payroll_recipient") == "payroll@company.example.com", "本周范围或工资收件人不一致")


def verify_workbook(path, checks):
    wb = load_workbook(path, data_only=False)
    expected_sheets = [
        "Source_Week", "Source_Convention", "Source_System_State", "Evidence_Register", "Scope_Decisions",
        "BambooHR_Drafts", "Jira_Drafts", "Payroll_Draft", "Tracker_Drafts", "Skipped_Employees",
        "Cross_System_Reconciliation", "Control_Summary",
    ]
    checks.require("R003", wb.sheetnames == expected_sheets, f"Excel 工作表应为 12 张，实际 {wb.sheetnames}")
    formulas = [cell.value for row in wb["Scope_Decisions"].iter_rows(min_row=2) for cell in row if isinstance(cell.value, str) and cell.value.startswith("=")]
    joined = "\n".join(formulas)
    checks.require("R004", len(formulas) == 36, f"Scope_Decisions 应有 36 个公式，实际 {len(formulas)}")
    for token in ("Source_Week!H4", "DATE(2026,3,18)", "DATE(2026,3,22)", "PROCESS_OFFBOARDING", "SKIP_OUT_OF_SCOPE", 'TEXT(E4,"yyyy-mm-dd")'):
        checks.require("R004", token in joined, f"公式缺少 {token}")

    reconcile = rows(wb["Cross_System_Reconciliation"])
    checks.require("R005", reconcile == [("Greg Foster", "EMP-2001", 1, 1, 1, 1, "YES"), ("Ian Costa", "EMP-2004", 1, 1, 1, 1, "YES")], "Greg/Ian 四系统对账不一致")
    skipped = rows(wb["Skipped_Employees"])
    checks.require("R006", len(skipped) == 1 and skipped[0][0:2] == ("Nadia Petrova", "EMP-2005") and skipped[0][2].strftime("%Y-%m-%d") == "2026-04-05" and skipped[0][3:] == ("Pending", "SKIP_OUT_OF_SCOPE", "Last day is after 2026-03-22"), "Nadia 跳过边界不正确")


def verify_artifacts(artifact_dir, checks):
    bamboo = json.loads((artifact_dir / BAMBOO).read_text())
    updates = bamboo.get("updates", [])
    expected_bamboo = [
        ("Greg Foster", "EMP-2001", "Pending", "Inactive", "2026-03-21", "DRAFT_NOT_APPLIED"),
        ("Ian Costa", "EMP-2004", "Pending", "Inactive", "2026-03-20", "DRAFT_NOT_APPLIED"),
    ]
    observed_bamboo = [(x.get("employee"), x.get("employee_id"), x.get("current_status"), x.get("target_status"), x.get("effective_date"), x.get("state")) for x in updates]
    checks.require("R007", bamboo.get("draft_count") == 2 and bamboo.get("applied_count") == 0 and observed_bamboo == expected_bamboo, "BambooHR 草稿应只包含 Greg/Ian 且未应用")

    jira = json.loads((artifact_dir / JIRA).read_text())
    tickets = jira.get("tickets", [])
    checks.require("R008", jira.get("draft_count") == 2 and jira.get("created_count") == 0 and len(tickets) == 2, "Jira 草稿数量或执行状态不正确")
    if len(tickets) == 2:
        for ticket, name, employee_id, date in zip(tickets, ("Greg Foster", "Ian Costa"), ("EMP-2001", "EMP-2004"), ("2026-03-21", "2026-03-20")):
            text = f"{ticket.get('summary', '')}\n{ticket.get('description', '')}"
            checks.require("R008", name in text and employee_id in text and date in text and "Revoke" in text and ticket.get("state") == "DRAFT_NOT_CREATED", f"Jira 草稿缺失 {name} 的身份、日期或离线状态")
    checks.require("R008", all("Nadia Petrova" not in json.dumps(ticket) and "EMP-2005" not in json.dumps(ticket) for ticket in tickets), "Jira 不得包含 Nadia")

    messages = list(mailbox.mbox(artifact_dir / PAYROLL, create=False))
    checks.require("R008", len(messages) == 1 and str(messages[0].get("To")) == "payroll@company.example.com", "工资通知应为 1 封离线草稿")
    if messages:
        body = body_text(messages[0])
        for token in ("2 in-scope employees", "Greg Foster", "EMP-2001", "2026-03-21", "5 days", "Ian Costa", "EMP-2004", "2026-03-20", "0 days", "DRAFT_NOT_SENT"):
            checks.require("R008", token in body, f"工资草稿缺少 {token}")
        checks.require("R008", "Nadia Petrova" in body and "intentionally excluded" in body and "3 days" not in body, "工资草稿必须说明 Nadia 被排除且不得带入其 PTO")

    tracker = json.loads((artifact_dir / TRACKER).read_text())
    tracker_rows = tracker.get("updates", [])
    observed_tracker = [(x.get("source_row"), x.get("employee"), x.get("employee_id"), x.get("previous_status"), x.get("target_status"), x.get("applied")) for x in tracker_rows]
    checks.require("R009", tracker.get("draft_count") == 2 and tracker.get("applied_count") == 0 and observed_tracker == [(2, "Greg Foster", "EMP-2001", "Pending", "Processed", False), (3, "Ian Costa", "EMP-2004", "Pending", "Processed", False)], "Tracker 草稿边界或 Processed 状态不正确")
    control = json.loads((artifact_dir / CONTROL).read_text())
    expected_control = {
        "employees_reviewed": 3, "employees_in_scope": 2, "employees_skipped": 1,
        "bamboo_drafts": 2, "bamboo_updates_applied": 0, "jira_drafts": 2, "jira_tickets_created": 0,
        "payroll_drafts": 1, "payroll_messages_sent": 0, "tracker_drafts": 2, "tracker_updates_applied": 0,
        "external_actions_executed": 0, "real_side_effects": 0,
    }
    checks.require("R009", all(control.get(key) == value for key, value in expected_control.items()), "执行控制计数或零副作用边界不正确")


def verify_office(artifact_dir, checks):
    text, sections, tables = docx_text(artifact_dir / DOCX)
    checks.require("R010", sections == 3 and tables == 3, f"Word 应为 3 页/3 表，实际 sections={sections}, tables={tables}")
    for token in ("Greg Foster", "Ian Costa", "Nadia Petrova", "EMP-2005", "2026-04-05", "SKIP_OUT_OF_SCOPE", "5 days", "0 days", "DRAFT_NOT_SENT"):
        checks.require("R010", token.lower() in text.lower(), f"Word 缺少 {token}")
    checks.require("R010", "Nadia Petrova" in text and "no BambooHR update" in text and "no payroll mention" in text, "Word 未明确 Nadia 的跨系统禁止动作")

    text, slides, table_counts, pictures = pptx_text(artifact_dir / PPTX)
    checks.require("R011", slides == 3 and table_counts == [1, 1, 1] and pictures == 0, f"PPT 应为 3 页且每页 1 表，实际 slides={slides}, tables={table_counts}, pictures={pictures}")
    for token in ("Greg Foster", "Ian Costa", "Nadia Petrova", "out of scope", "BambooHR", "Jira", "Payroll", "Tracker", "DRAFT_NOT_APPLIED", "DRAFT_NOT_SENT", "0 real side effects"):
        checks.require("R011", token.lower() in text.lower(), f"PPT 缺少 {token}")

    for name in (XLSX, DOCX, PPTX):
        safe, detail = package_safe(artifact_dir / name)
        checks.require("R012", safe, f"{name} OOXML 安全检查失败：{detail}")


def finish(checks, render_pages):
    hurdles = [key for key in CAPS if not checks.ok(key)]
    base = sum(weight for key, weight in WEIGHTS.items() if checks.ok(key)) / 100
    cap = min([1] + [CAPS[key] for key in hurdles])
    penalty = (
        (0.30 if any(not checks.ok(key) for key in ("R002", "R003", "R004", "R005", "R006")) else 0)
        + (0.30 if any(not checks.ok(key) for key in ("R007", "R008", "R009")) else 0)
        + (0.25 if any(not checks.ok(key) for key in ("R001", "R010", "R011", "R012")) else 0)
    )
    reward = round(max(0, min(base, cap) - penalty), 6)
    return {
        "task_id": TASK,
        "reward": reward,
        "passed": reward >= 0.8 and not hurdles,
        "criteria": checks.values,
        "hurdle_failures": hurdles,
        "source_truth": {"reviewed": 3, "in_scope": 2, "skipped": 1, "external_actions": 0},
        "render_pages": render_pages,
    }


def evaluate(artifact_dir):
    checks = Checks()
    verify_source(checks)
    files = {path.name for path in artifact_dir.iterdir() if path.is_file()} if artifact_dir.is_dir() else set()
    checks.require("R001", files == REQUIRED, f"交付集合不一致：{files}")
    if files != REQUIRED:
        return finish(checks, {})

    for function, criteria in (
        (lambda: verify_workbook(artifact_dir / XLSX, checks), ("R003", "R004", "R005", "R006", "R012")),
        (lambda: verify_artifacts(artifact_dir, checks), ("R007", "R008", "R009")),
        (lambda: verify_office(artifact_dir, checks), ("R010", "R011", "R012")),
    ):
        try:
            function()
        except Exception as exc:
            for criterion in criteria:
                checks.require(criterion, False, str(exc))

    pages = {}
    for name, expected in ((XLSX, 12), (DOCX, 3), (PPTX, 3)):
        try:
            count, rendered_text = render(artifact_dir / name)
            pages[name] = count
            checks.require("R012", count == expected, f"{name} 应渲染 {expected} 页，实际 {count}")
            checks.require("R012", not re.search(r"#(?:VALUE|REF|NAME|DIV/0|N/A)", rendered_text, re.I), f"{name} 出现公式错误")
        except Exception as exc:
            checks.require("R012", False, str(exc))
    return finish(checks, pages)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n")
    args.reward_file.write_text(f"{result['reward']:.6f}\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
