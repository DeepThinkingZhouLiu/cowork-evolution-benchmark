#!/usr/bin/env python3
"""AutomationBench 5070 多系统离职边界、动作草稿与跨系统对账逐题解法。"""

import argparse
import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from office_runtime import GREEN, write_docx, write_mbox, write_pptx, write_workbook


INPUT_HASHES = {
    "Initial_Application_State.json": "ae2fb67332f1613150faaf09df20b906be87de8609382b7d4482fcc1e3a1e1dc",
    "Offboarding_Master.xlsx": "bf25180c05172e57d11eefc87a0f522ea18384f645b3426f857e641d548dd742",
    "Workflow_Context.json": "a11d79dcd287ae48fed36f167220293cf0a95200000012e2135aa9b9abe5e476",
}
OUTPUTS = (
    "Offboarding_Multi_System_Audit.xlsx",
    "Offboarding_Operations_Brief.docx",
    "Offboarding_Operations_Review.pptx",
    "BambooHR_Status_Update_Drafts.json",
    "Jira_Access_Revocation_Drafts.json",
    "Payroll_Notification_Draft.mbox",
    "Tracker_Status_Update_Drafts.json",
    "Offboarding_Execution_Control.json",
)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_truth(input_dir: Path):
    for name, expected in INPUT_HASHES.items():
        path = input_dir / name
        if not path.is_file() or sha256(path) != expected:
            raise ValueError(f"输入 {name} 缺失或哈希不一致")

    wb = load_workbook(input_dir / "Offboarding_Master.xlsx", read_only=True, data_only=True)
    ws = wb["Week of March 18"]
    headers = [cell.value for cell in ws[1]]
    employees = [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]
    for employee in employees:
        if isinstance(employee["Last Day"], datetime):
            employee["Last Day"] = employee["Last Day"].date()

    truth = [
        (2, "Greg Foster", "EMP-2001", "greg.foster@company.example.com", "2026-03-21", "Voluntary", "5 days", "Pending"),
        (3, "Ian Costa", "EMP-2004", "ian.costa@company.example.com", "2026-03-20", "Voluntary", "0 days", "Pending"),
        (4, "Nadia Petrova", "EMP-2005", "nadia.petrova@company.example.com", "2026-04-05", "Voluntary", "3 days", "Pending"),
    ]
    observed = [
        (
            row["Source Row"],
            row["Employee"],
            row["Employee ID"],
            row["Email"],
            row["Last Day"].isoformat(),
            row["Type"],
            row["PTO Remaining"],
            row["Status"],
        )
        for row in employees
    ]
    if observed != truth:
        raise ValueError("5070 冻结人员清单与人工真值不一致")

    state = json.loads((input_dir / "Initial_Application_State.json").read_text())
    context = json.loads((input_dir / "Workflow_Context.json").read_text())
    if state != {"bamboohr": {"actions": {}}, "jira": {"actions": {}}}:
        raise ValueError("5070 初始应用状态不一致")
    if (context["scope_start"], context["scope_end"], context["payroll_recipient"]) != (
        "2026-03-18",
        "2026-03-22",
        "payroll@company.example.com",
    ):
        raise ValueError("5070 工作流边界不一致")
    return employees, state, context


def decision_formulas():
    formulas = []
    for row in range(2, 5):
        formulas.append(
            [f"=Source_Week!{column}{row}" for column in "ABCDEFGH"]
            + [
                f'=IF(AND(E{row}>=DATE(2026,3,18),E{row}<=DATE(2026,3,22)),"YES","NO")',
                f'=IF(I{row}="YES","PROCESS_OFFBOARDING","SKIP_OUT_OF_SCOPE")',
                f'=IF(I{row}="YES","Processed",H{row})',
                f'="employee_id="&C{row}&"|last_day="&TEXT(E{row},"yyyy-mm-dd")&"|pto="&G{row}&"|decision="&J{row}&"|target="&K{row}',
            ]
        )
    return formulas


def build(output_dir: Path, employees, state, context):
    in_scope = employees[:2]
    skipped = employees[2:]
    bamboo = [
        {
            "employee": row["Employee"],
            "employee_id": row["Employee ID"],
            "current_status": row["Status"],
            "target_status": "Inactive",
            "effective_date": row["Last Day"].isoformat(),
            "state": "DRAFT_NOT_APPLIED",
        }
        for row in in_scope
    ]
    jira = [
        {
            "summary": f"Revoke access for {row['Employee']} ({row['Employee ID']})",
            "description": (
                f"Revoke company-system access for {row['Employee']} ({row['Employee ID']}) "
                f"after the last day {row['Last Day'].isoformat()}. Offboarding type: {row['Type']}."
            ),
            "state": "DRAFT_NOT_CREATED",
        }
        for row in in_scope
    ]
    tracker = [
        {
            "source_row": row["Source Row"],
            "employee": row["Employee"],
            "employee_id": row["Employee ID"],
            "previous_status": row["Status"],
            "target_status": "Processed",
            "applied": False,
        }
        for row in in_scope
    ]

    sheets = [
        {
            "name": "Source_Week",
            "headers": ["Row", "Employee", "Employee ID", "Email", "Last Day", "Type", "PTO Remaining", "Status"],
            "rows": [
                [row[key] for key in ("Source Row", "Employee", "Employee ID", "Email", "Last Day", "Type", "PTO Remaining", "Status")]
                for row in employees
            ],
            "widths": [10, 26, 18, 42, 18, 18, 22, 18],
            "table": "SourceWeek",
        },
        {
            "name": "Source_Convention",
            "headers": ["Row", "Outcome", "Tracker Status"],
            "rows": [[2, "All required offboarding actions completed", "Processed"]],
            "widths": [10, 64, 24],
            "table": "SourceConvention",
        },
        {
            "name": "Source_System_State",
            "headers": ["System", "Initial State", "Action Boundary"],
            "rows": [["BambooHR", "No actions", "Draft only"], ["Jira", "No actions", "Draft only"]],
            "widths": [24, 32, 44],
            "table": "SourceSystemState",
        },
        {
            "name": "Evidence_Register",
            "headers": ["Evidence", "Supports", "SHA256"],
            "rows": [[name, "Frozen AutomationBench 5070 source", digest] for name, digest in INPUT_HASHES.items()],
            "widths": [62, 54, 68],
            "table": "EvidenceRegister",
        },
        {
            "name": "Scope_Decisions",
            "headers": [
                "Row", "Employee", "Employee ID", "Email", "Last Day", "Type", "PTO Remaining", "Source Status",
                "In Scope", "Decision", "Target Status", "Trace",
            ],
            "rows": decision_formulas(),
            "widths": [10, 26, 18, 42, 18, 18, 22, 18, 16, 26, 20, 112],
            "table": "ScopeDecisions",
            "conditional": [{"range": "I2:I4", "formula": 'I2="YES"', "color": GREEN}],
        },
        {
            "name": "BambooHR_Drafts",
            "headers": ["Employee", "Employee ID", "Current", "Target", "Effective Date", "State"],
            "rows": [[x[key] for key in ("employee", "employee_id", "current_status", "target_status", "effective_date", "state")] for x in bamboo],
            "widths": [26, 18, 18, 18, 20, 28],
            "table": "BambooDrafts",
        },
        {
            "name": "Jira_Drafts",
            "headers": ["Summary", "Description", "State"],
            "rows": [[x["summary"], x["description"], x["state"]] for x in jira],
            "widths": [54, 116, 28],
            "table": "JiraDrafts",
        },
        {
            "name": "Payroll_Draft",
            "headers": ["Recipient", "Employees", "Greg PTO", "Ian PTO", "State"],
            "rows": [[context["payroll_recipient"], "Greg Foster; Ian Costa", "5 days", "0 days", "DRAFT_NOT_SENT"]],
            "widths": [42, 44, 18, 18, 26],
            "table": "PayrollDraft",
        },
        {
            "name": "Tracker_Drafts",
            "headers": ["Row", "Employee", "Employee ID", "Previous", "Target", "Applied"],
            "rows": [[x[key] for key in ("source_row", "employee", "employee_id", "previous_status", "target_status", "applied")] for x in tracker],
            "widths": [10, 26, 18, 18, 18, 16],
            "table": "TrackerDrafts",
        },
        {
            "name": "Skipped_Employees",
            "headers": ["Employee", "Employee ID", "Last Day", "Source Status", "Decision", "Reason"],
            "rows": [[row["Employee"], row["Employee ID"], row["Last Day"], row["Status"], "SKIP_OUT_OF_SCOPE", "Last day is after 2026-03-22"] for row in skipped],
            "widths": [28, 18, 18, 20, 26, 52],
            "table": "SkippedEmployees",
        },
        {
            "name": "Cross_System_Reconciliation",
            "headers": ["Employee", "Employee ID", "Bamboo", "Jira", "Payroll", "Tracker", "Matched"],
            "rows": [[row["Employee"], row["Employee ID"], 1, 1, 1, 1, "YES"] for row in in_scope],
            "widths": [26, 18, 16, 16, 16, 16, 18],
            "table": "CrossSystemReconciliation",
        },
        {
            "name": "Control_Summary",
            "headers": ["Metric", "Count", "State"],
            "rows": [
                ["Employees reviewed", 3, "Greg; Ian; Nadia"],
                ["Employees in scope", 2, "Greg; Ian"],
                ["Employees skipped", 1, "Nadia: out of scope"],
                ["BambooHR drafts", 2, "NOT_APPLIED"],
                ["Jira drafts", 2, "NOT_CREATED"],
                ["Payroll drafts", 1, "NOT_SENT"],
                ["Tracker drafts", 2, "NOT_APPLIED"],
                ["External actions executed", 0, "NONE"],
                ["Real side effects", 0, "NONE"],
            ],
            "widths": [48, 16, 76],
            "table": "ControlSummary",
            "chart": {"title": "Offboarding draft coverage", "label_col": 1, "value_col": 2, "min_row": 4, "max_row": 8},
        },
    ]
    write_workbook(output_dir / OUTPUTS[0], "Multi-System Offboarding Audit", sheets)

    doc_pages = [
        {
            "blocks": [
                {"type": "heading", "text": "This-week scope is exactly two employees"},
                {"type": "paragraph", "text": "The frozen window is 2026-03-18 through 2026-03-22. Greg Foster and Ian Costa are in scope; Nadia Petrova is outside the window and remains Pending."},
                {"type": "table", "headers": ["Employee", "ID", "Last day", "PTO", "Decision"], "rows": [[r["Employee"], r["Employee ID"], r["Last Day"].isoformat(), r["PTO Remaining"], "PROCESS_OFFBOARDING"] for r in in_scope]},
            ]
        },
        {
            "blocks": [
                {"type": "heading", "text": "Four-system draft reconciliation"},
                {"type": "table", "headers": ["Employee", "BambooHR", "Jira", "Payroll", "Tracker"], "rows": [[r["Employee"], "1 draft", "1 draft", "included", "Processed draft"] for r in in_scope]},
                {"type": "paragraph", "text": "Greg Foster has 5 days PTO remaining; Ian Costa has 0 days. Both values are preserved in the payroll notification draft."},
            ]
        },
        {
            "blocks": [
                {"type": "heading", "text": "Skip boundary and offline safety"},
                {"type": "paragraph", "text": "Nadia Petrova (EMP-2005), last day 2026-04-05, is SKIP_OUT_OF_SCOPE: no BambooHR update, no Jira ticket, no payroll mention, and no tracker update."},
                {"type": "table", "headers": ["Action", "Prepared", "Executed", "State"], "rows": [["BambooHR", 2, 0, "DRAFT_NOT_APPLIED"], ["Jira", 2, 0, "DRAFT_NOT_CREATED"], ["Payroll", 1, 0, "DRAFT_NOT_SENT"], ["Tracker", 2, 0, "NOT_APPLIED"], ["Real side effects", 0, 0, "NONE"]]},
            ]
        },
    ]
    write_docx(output_dir / OUTPUTS[1], "Multi-System Offboarding Operations Brief", "Week of 2026-03-18 · exact scope · offline drafts only", doc_pages)

    slides = [
        {
            "title": "Two employees are in this-week scope",
            "subtitle": "Window 2026-03-18 to 2026-03-22 · Nadia Petrova is out of scope",
            "headers": ["Employee", "ID", "Last day", "PTO", "Decision"],
            "rows": [[r["Employee"], r["Employee ID"], r["Last Day"].isoformat(), r["PTO Remaining"], "PROCESS"] for r in in_scope],
        },
        {
            "title": "Four systems reconcile one-to-one",
            "subtitle": "BambooHR, Jira, payroll, and tracker are prepared for Greg and Ian only",
            "headers": ["Employee", "BambooHR", "Jira", "Payroll", "Tracker"],
            "rows": [[r["Employee"], "1 draft", "1 draft", "included", "Processed"] for r in in_scope],
        },
        {
            "title": "Nadia is skipped; every action remains offline",
            "subtitle": "0 external actions · 0 real side effects",
            "headers": ["Action", "Prepared", "Executed", "State"],
            "rows": [["BambooHR", 2, 0, "DRAFT_NOT_APPLIED"], ["Jira", 2, 0, "DRAFT_NOT_CREATED"], ["Payroll", 1, 0, "DRAFT_NOT_SENT"], ["Tracker", 2, 0, "NOT_APPLIED"]],
        },
    ]
    write_pptx(output_dir / OUTPUTS[2], slides)

    (output_dir / OUTPUTS[3]).write_text(json.dumps({"task_id": "automationbench-5070", "draft_count": 2, "applied_count": 0, "updates": bamboo}, ensure_ascii=False, indent=2) + "\n")
    (output_dir / OUTPUTS[4]).write_text(json.dumps({"task_id": "automationbench-5070", "draft_count": 2, "created_count": 0, "tickets": jira}, ensure_ascii=False, indent=2) + "\n")
    write_mbox(
        output_dir / OUTPUTS[5],
        [
            {
                "to": context["payroll_recipient"],
                "subject": "Offboarding payroll notification: Greg Foster and Ian Costa",
                "type": "OFFBOARDING_PAYROLL_NOTIFICATION",
                "id_prefix": "offboarding-payroll-5070",
                "body": (
                    "Please prepare final payroll for 2 in-scope employees.\n\n"
                    "Greg Foster | EMP-2001 | last day 2026-03-21 | Voluntary | PTO remaining 5 days\n"
                    "Ian Costa | EMP-2004 | last day 2026-03-20 | Voluntary | PTO remaining 0 days\n\n"
                    "Nadia Petrova is outside this-week scope and is intentionally excluded.\n"
                    "DRAFT_NOT_SENT."
                ),
            }
        ],
    )
    (output_dir / OUTPUTS[6]).write_text(json.dumps({"task_id": "automationbench-5070", "draft_count": 2, "applied_count": 0, "updates": tracker}, ensure_ascii=False, indent=2) + "\n")
    control = {
        "task_id": "automationbench-5070",
        "employees_reviewed": 3,
        "employees_in_scope": 2,
        "employees_skipped": 1,
        "bamboo_drafts": 2,
        "bamboo_updates_applied": 0,
        "jira_drafts": 2,
        "jira_tickets_created": 0,
        "payroll_drafts": 1,
        "payroll_messages_sent": 0,
        "tracker_drafts": 2,
        "tracker_updates_applied": 0,
        "external_actions_executed": 0,
        "real_side_effects": 0,
    }
    (output_dir / OUTPUTS[7]).write_text(json.dumps(control, ensure_ascii=False, indent=2) + "\n")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    employees, state, context = read_truth(args.input_dir)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    for child in args.output_dir.iterdir():
        if child.is_file():
            child.unlink()
        elif child.is_dir():
            shutil.rmtree(child)
    build(args.output_dir, employees, state, context)
    if {item.name for item in args.output_dir.iterdir()} != set(OUTPUTS):
        raise ValueError("交付文件集合不一致")


if __name__ == "__main__":
    main()
