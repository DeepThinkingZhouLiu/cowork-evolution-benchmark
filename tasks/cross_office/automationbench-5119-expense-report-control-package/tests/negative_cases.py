#!/usr/bin/env python3
"""向 AutomationBench 5119 Golden 逐项注入本题专属错误，验证 Judge 拒绝坏答案。"""

from __future__ import annotations

import argparse
import importlib.util
import json
import mailbox
import shutil
import sys
import tempfile
import zipfile
from email.generator import BytesGenerator
from email.message import EmailMessage
from email.policy import SMTP
from io import BytesIO
from pathlib import Path
from typing import Callable

from docx import Document
from lxml import etree
from openpyxl import load_workbook
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
GOLDEN = ROOT / "solution" / "Golden"
XLSX = "Expense_Report_Audit.xlsx"
DOCX = "Expense_Control_Brief.docx"
PPTX = "Expense_Review.pptx"
MBOX = "Expense_Notification_Drafts.mbox"
UPDATES = "Expense_Tracker_Update_Drafts.json"
SLACK = "Expense_Slack_Control.json"
EXECUTION = "Expense_Execution_Control.json"
FILES = (XLSX, DOCX, PPTX, MBOX, UPDATES, SLACK, EXECUTION)
Mutator = Callable[[Path], None]
Case = tuple[str, Mutator, str]


def load_judge():
    spec = importlib.util.spec_from_file_location("expense_5119_judge", ROOT / "tests" / "judge.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("无法载入本题 Judge")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    module.render = lambda path: ({".xlsx": 10, ".docx": 3, ".pptx": 3}[path.suffix], "render-ok")
    return module


def rewrite_zip(path: Path, transform: Callable[[str, bytes], bytes | None]) -> None:
    temporary = path.with_suffix(path.suffix + ".negative")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = transform(info.filename, source.read(info.filename))
            if data is not None:
                target.writestr(info, data)
    temporary.replace(path)


def sheet_xml_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = etree.fromstring(archive.read("xl/workbook.xml"))
    relationships = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    ns = {
        "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    }
    relationship_id = workbook.xpath(f'//m:sheet[@name="{sheet_name}"]/@r:id', namespaces=ns)[0]
    target = relationships.xpath(f'//pr:Relationship[@Id="{relationship_id}"]/@Target', namespaces=ns)[0]
    return target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")


def edit_cell(path: Path, sheet_name: str, coordinate: str, *, value=None, remove_formula: bool = False) -> None:
    with zipfile.ZipFile(path, "r") as archive:
        target_sheet = sheet_xml_path(archive, sheet_name)
    changed = False
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    def transform(name: str, data: bytes) -> bytes:
        nonlocal changed
        if name != target_sheet:
            return data
        root = etree.fromstring(data)
        cells = root.xpath(f'//m:c[@r="{coordinate}"]', namespaces={"m": namespace})
        if not cells:
            raise ValueError(f"{sheet_name}!{coordinate} 不存在")
        cell = cells[0]
        if remove_formula:
            node = cell.find(f"{{{namespace}}}f")
            if node is None:
                raise ValueError(f"{sheet_name}!{coordinate} 没有公式")
            cell.remove(node)
        if value is not None:
            inline = cell.find(f"{{{namespace}}}is")
            if inline is not None:
                cell.remove(inline)
            node = cell.find(f"{{{namespace}}}v")
            if node is None:
                node = etree.SubElement(cell, f"{{{namespace}}}v")
            if isinstance(value, bool):
                node.text = "1" if value else "0"
                cell.set("t", "b")
            elif isinstance(value, (int, float)):
                node.text = str(value)
                cell.attrib.pop("t", None)
            else:
                node.text = str(value)
                cell.set("t", "str")
        changed = True
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    rewrite_zip(path, transform)
    if not changed:
        raise ValueError("目标单元格未修改")


def openpyxl_mutation(path: Path, mutator: Callable) -> None:
    workbook = load_workbook(path)
    mutator(workbook)
    workbook.save(path)


def mutate_json(path: Path, mutator: Callable[[dict], None]) -> None:
    payload = json.loads(path.read_text(encoding="utf-8"))
    mutator(payload)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def decode_message(message) -> str:
    if message.is_multipart():
        return "\n".join((part.get_payload(decode=True) or b"").decode(part.get_content_charset() or "utf-8", errors="replace") for part in message.walk() if part.get_content_type() == "text/plain").strip()
    return (message.get_payload(decode=True) or b"").decode(message.get_content_charset() or "utf-8", errors="replace").strip()


def mbox_states(path: Path) -> list[dict]:
    box = mailbox.mbox(path, create=False)
    messages = list(box)
    box.close()
    return [
        {
            "from": str(message.get("From", "")), "to": str(message.get("To", "")), "cc": str(message.get("Cc", "")), "bcc": str(message.get("Bcc", "")),
            "subject": str(message.get("Subject", "")), "date": str(message.get("Date", "")), "message_id": str(message.get("Message-ID", "")),
            "status": str(message.get("X-EvoCowork-Status", "")), "source_row": str(message.get("X-EvoCowork-Source-Row", "")),
            "decision": str(message.get("X-EvoCowork-Decision", "")), "proposed_status": str(message.get("X-EvoCowork-Proposed-Status", "")),
            "order": str(message.get("X-EvoCowork-Order", "")), "body": decode_message(message),
        }
        for message in messages
    ]


def write_mbox_states(path: Path, states: list[dict]) -> None:
    chunks = []
    for order, state in enumerate(states, 1):
        message = EmailMessage(policy=SMTP)
        for key, header in (
            ("from", "From"), ("to", "To"), ("cc", "Cc"), ("bcc", "Bcc"), ("subject", "Subject"), ("date", "Date"),
            ("message_id", "Message-ID"), ("status", "X-EvoCowork-Status"), ("source_row", "X-EvoCowork-Source-Row"),
            ("decision", "X-EvoCowork-Decision"), ("proposed_status", "X-EvoCowork-Proposed-Status"), ("order", "X-EvoCowork-Order"),
        ):
            if state.get(key):
                message[header] = state[key]
        message.set_content(state["body"].rstrip() + "\n", charset="utf-8")
        buffer = BytesIO()
        BytesGenerator(buffer, policy=SMTP, mangle_from_=True, maxheaderlen=0).flatten(message)
        chunks.extend([f"From MAILER-DAEMON Wed Apr 15 09:00:{order:02d} 2026\n".encode("ascii"), buffer.getvalue().replace(b"\r\n", b"\n"), b"\n"])
    path.write_bytes(b"".join(chunks))


def mutate_message(path: Path, index: int, mutator: Callable[[dict], None]) -> None:
    states = mbox_states(path)
    mutator(states[index])
    write_mbox_states(path, states)


def replace_docx_text(path: Path, old: str, new: str, *, row_token: str | None = None) -> None:
    document = Document(path)
    changed = False
    if row_token is None:
        for paragraph in document.paragraphs:
            if old in paragraph.text:
                paragraph.text = paragraph.text.replace(old, new, 1)
                changed = True
                break
    if not changed:
        for table in document.tables:
            for row in table.rows:
                if row_token and row_token not in " | ".join(cell.text for cell in row.cells):
                    continue
                for cell in row.cells:
                    if old in cell.text:
                        cell.text = cell.text.replace(old, new, 1)
                        changed = True
                        break
                if changed:
                    break
            if changed:
                break
    if not changed:
        raise ValueError(f"Word 找不到 {old!r}")
    document.save(path)


def replace_pptx_text(path: Path, old: str, new: str) -> None:
    presentation = Presentation(path)
    changed = False
    for slide in presentation.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and old in shape.text:
                shape.text = shape.text.replace(old, new, 1)
                changed = True
                break
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        if old in cell.text:
                            cell.text = cell.text.replace(old, new, 1)
                            changed = True
                            break
                    if changed:
                        break
            if changed:
                break
        if changed:
            break
    if not changed:
        raise ValueError(f"PPT 找不到 {old!r}")
    presentation.save(path)


def replace_pptx_row_text(path: Path, row_token: str, old: str, new: str) -> None:
    presentation = Presentation(path)
    changed = False
    for slide in presentation.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue
            for row in shape.table.rows:
                if row_token not in " | ".join(cell.text for cell in row.cells):
                    continue
                for cell in row.cells:
                    if cell.text == old:
                        cell.text = new
                        changed = True
                        break
                if changed:
                    break
            if changed:
                break
        if changed:
            break
    if not changed:
        raise ValueError(f"PPT 在 {row_token!r} 行找不到 {old!r}")
    presentation.save(path)


def drop_ppt_slide(path: Path) -> None:
    presentation = Presentation(path)
    presentation.slides._sldIdLst.remove(presentation.slides._sldIdLst[-1])
    presentation.save(path)


def build_cases() -> list[Case]:
    cases: list[Case] = []
    for name in FILES:
        cases.append((f"missing-{name}", lambda root, n=name: (root / n).unlink(), "R001"))
    cases.append(("extra-file", lambda root: (root / "unexpected.txt").write_text("unexpected", encoding="utf-8"), "R001"))
    corrupt_criterion = {XLSX:"R001", DOCX:"R001", PPTX:"R001", MBOX:"R012", UPDATES:"R013", SLACK:"R014", EXECUTION:"R014"}
    for name in FILES:
        cases.append((f"corrupt-{name}", lambda root, n=name: (root / n).write_bytes(b"not-a-valid-artifact"), corrupt_criterion[name]))

    source_wrong = {"A":99,"B":"Wrong Person","C":"wrong@company.example.com","D":"Other","E":"$9,999.00","F":"Wrong description","G":"MISSING","H":"Approved","I":"Wrong note"}
    for row in range(2, 10):
        for column, value in source_wrong.items():
            cases.append((f"source-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Source_Reports", c, value=v), "R002"))

    for column, value in (("A","wrong-id"),("B","wrong@company.example.com"),("C","wrong@company.example.com"),("D","Wrong policy"),("E","Thu, 16 Apr 2026 09:00:00 +0000"),("F","Receipts are optional below $100."),("G","wrong-source")):
        cases.append((f"policy-{column}", lambda root, c=f"{column}2", v=value: edit_cell(root / XLSX, "Policy_Source", c, value=v), "R003"))
    for row in range(2, 21):
        for column, value in (("A","WRONG-ID"),("D","Wrong fact"),("E","wrong-source")):
            cases.append((f"evidence-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Evidence_Register", c, value=v), "R003"))

    for column in range(1, 23):
        letter = chr(64 + column)
        cases.append((f"formula-{letter}2", lambda root, c=f"{letter}2": edit_cell(root / XLSX, "Expense_Decisions", c, remove_formula=True), "R004"))

    people = {2:"R005",3:"R006",4:"R007",5:"R005",6:"R010",7:"R008",8:"R009",9:"R005"}
    decision_wrong = {"F":9999,"K":"WRONG","L":"WRONG","M":"WRONG","N":"Wrong Role","O":"WRONG_DECISION","P":"Wrong Status","Q":"wrong@company.example.com","R":"WRONG","S":"WRONG","T":"WRONG","U":"Wrong authority","V":"wrong-source"}
    for row, criterion in people.items():
        for column, value in decision_wrong.items():
            cases.append((f"decision-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Expense_Decisions", c, value=v), criterion))

    approved_criteria = {2:"R005",3:"R005",4:"R009",5:"R005"}
    for row, criterion in approved_criteria.items():
        for column, value in (("C","Wrong Person"),("E",9999),("F","No"),("G","WRONG"),("H","wrong@company.example.com"),("I","Wrong Status"),("J","Wrong basis"),("K","wrong-source")):
            cases.append((f"approved-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Approved_Queue", c, value=v), criterion))
    higher_criteria = {2:"R006",3:"R007",4:"R008"}
    for row, criterion in higher_criteria.items():
        for column, value in (("C","Wrong Person"),("E",9999),("F","Wrong description"),("G","No"),("H","Wrong Role"),("I","wrong@company.example.com"),("J","Wrong Status"),("K","Wrong basis"),("L","wrong-source")):
            cases.append((f"higher-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Higher_Approval_Queue", c, value=v), criterion))
    for column, value in (("B","Wrong Elena"),("D",99),("E","Yes"),("F","Approved"),("G","APPROVED"),("H","Approved"),("I","YES"),("J","YES"),("K","YES"),("L","Receipt bypass allowed"),("M","wrong-source")):
        cases.append((f"silent-{column}", lambda root, c=f"{column}2", v=value: edit_cell(root / XLSX, "Silent_Hold_Control", c, value=v), "R010"))

    for row in range(2, 9):
        for column, value in (("C","Wrong Person"),("D","WRONG"),("E","wrong@company.example.com"),("F","cc@company.example.com"),("G","bcc@company.example.com"),("H","Wrong subject"),("I","SENT"),("J","wrong-source")):
            cases.append((f"notification-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Notification_Register", c, value=v), "R012"))
    for row in range(2, 9):
        for column, value in (("C","Wrong Person"),("D","Approved"),("E","WRONG"),("F","Wrong Status"),("G",True),("H","APPLIED"),("I","wrong-source")):
            cases.append((f"tracker-{column}{row}", lambda root, c=f"{column}{row}", v=value: edit_cell(root / XLSX, "Tracker_Update_Register", c, value=v), "R013"))

    for row in range(2, 19):
        criterion = "R014" if row in (13, 15, 16, 17, 18) else "R011"
        cases.append((f"summary-B{row}", lambda root, c=f"B{row}": edit_cell(root / XLSX, "Control_Summary", c, value=99), criterion))
    cases.extend([
        ("xlsx-table-removed", lambda root: openpyxl_mutation(root / XLSX, lambda wb: wb["Source_Reports"].tables.clear()), "R018"),
        ("xlsx-chart-removed", lambda root: openpyxl_mutation(root / XLSX, lambda wb: setattr(wb["Control_Summary"], "_charts", [])), "R018"),
        ("xlsx-cf-removed", lambda root: openpyxl_mutation(root / XLSX, lambda wb: setattr(wb["Expense_Decisions"], "conditional_formatting", type(wb["Expense_Decisions"].conditional_formatting)())), "R018"),
    ])

    cases.extend([
        ("mbox-empty", lambda root: write_mbox_states(root / MBOX, []), "R012"),
        ("mbox-drop-alice", lambda root: write_mbox_states(root / MBOX, mbox_states(root / MBOX)[1:]), "R012"),
        ("mbox-drop-hector", lambda root: write_mbox_states(root / MBOX, mbox_states(root / MBOX)[:-1]), "R012"),
        ("mbox-duplicate", lambda root: write_mbox_states(root / MBOX, mbox_states(root / MBOX) + [mbox_states(root / MBOX)[-1]]), "R012"),
        ("mbox-reverse", lambda root: write_mbox_states(root / MBOX, list(reversed(mbox_states(root / MBOX)))), "R012"),
    ])
    action_criteria = ["R005","R006","R007","R005","R008","R009","R005"]
    for index, criterion in enumerate(action_criteria):
        for field, value in (("to","wrong@company.example.com"),("cc","cc@company.example.com"),("bcc","bcc@company.example.com"),("subject","Wrong subject"),("date","Thu, 16 Apr 2026 09:00:00 +0000"),("message_id","<wrong@evocowork.local>"),("status","SENT"),("source_row","99"),("decision","WRONG"),("proposed_status","Wrong Status"),("order","99")):
            cases.append((f"mbox-{index}-{field}", lambda root, i=index, f=field, v=value: mutate_message(root / MBOX, i, lambda state: state.update({f:v})), criterion))
    body_mutations = [
        (0,"$62.00","$162.00","R005"),(0,"is approved","requires CFO approval","R005"),
        (1,"$120.00","$75.00","R006"),(1,"Pending CFO Approval","Approved","R006"),
        (2,"$2,400.00","$240.00","R007"),(2,"business class","economy class","R007"),(2,"Pending VP Approval","Approved","R007"),
        (3,"$350.00","$550.00","R005"),(4,"$799.00","$499.00","R008"),(4,"Pending IT Director Approval","Approved","R008"),
        (5,"CFO pre-approval on April 3","no pre-approval","R009"),(6,"Economy flight to Chicago","Business class flight to Chicago","R005"),
        (0,"DRAFT_NOT_SENT","SENT","R005"),
    ]
    for index, old, new, criterion in body_mutations:
        def mutate_body(root, i=index, o=old, n=new):
            def change(state):
                if o not in state["body"]:
                    raise ValueError(f"邮件正文找不到 {o!r}")
                state["body"] = state["body"].replace(o, n, 1)
            mutate_message(root / MBOX, i, change)
        cases.append((f"mbox-body-{index}-{old[:10]}", mutate_body, criterion))

    update_top = [("task_id","wrong"),("status","APPLIED"),("frozen_time","2026-04-16T09:00:00Z"),("receipt_policy","Receipts optional below $100."),("draft_count",6),("applied_count",1),("source_modified",True)]
    for key, value in update_top:
        cases.append((f"updates-{key}", lambda root, k=key, v=value: mutate_json(root / UPDATES, lambda payload: payload.update({k:v})), "R013"))
    cases.extend([
        ("updates-precedence", lambda root: mutate_json(root / UPDATES, lambda payload: payload["precedence"].reverse()), "R013"),
        ("updates-drop-alice", lambda root: mutate_json(root / UPDATES, lambda payload: payload["drafts"].pop(0)), "R013"),
        ("updates-drop-hector", lambda root: mutate_json(root / UPDATES, lambda payload: payload["drafts"].pop()), "R013"),
        ("updates-reverse", lambda root: mutate_json(root / UPDATES, lambda payload: payload["drafts"].reverse()), "R013"),
        ("updates-add-elena", lambda root: mutate_json(root / UPDATES, lambda payload: payload["drafts"].append({"employee":"Elena Marchetti"})), "R010"),
    ])
    draft_wrong = {"source_row":99,"employee":"Wrong Person","employee_email":"wrong@company.example.com","category":"Other","amount":"$9,999.00","description":"Wrong","receipt":"MISSING","preapproved":"WRONG","decision":"WRONG","approval_role":"Wrong","notify_to":"wrong@company.example.com","source_status":"Approved","proposed_status":"Wrong","state":"APPLIED","applied":True,"source_trace":[]}
    for index, criterion in enumerate(action_criteria):
        for key, value in draft_wrong.items():
            cases.append((f"updates-{index}-{key}", lambda root, i=index, k=key, v=value: mutate_json(root / UPDATES, lambda payload: payload["drafts"][i].update({k:v})), criterion))

    slack_mutations = [("task_id","wrong"),("decision","POST"),("reason","Slack authorized"),("draft_count",1),("posted_count",1),("posted",True),("employee_names_in_slack",["Elena Marchetti"]),("expense_details_in_slack",["$1,800.00"]),("receipt_details_in_slack",["No"]),("messages_created",[{"text":"expense"}]),("source_trace",[])]
    for key, value in slack_mutations:
        cases.append((f"slack-{key}", lambda root, k=key, v=value: mutate_json(root / SLACK, lambda payload: payload.update({k:v})), "R014"))
    cases.extend([
        ("slack-channel-id", lambda root: mutate_json(root / SLACK, lambda payload: payload["source_channel"].update({"id":"C_OTHER"})), "R014"),
        ("slack-channel-private", lambda root: mutate_json(root / SLACK, lambda payload: payload["source_channel"].update({"is_private":True})), "R014"),
    ])

    execution_mutations = [
        ("task-id", lambda payload: payload.update({"task_id":"wrong"})), ("status", lambda payload: payload.update({"status":"EXECUTED"})), ("side-effect", lambda payload: payload.update({"real_side_effect_count":1})),
        ("email-approved", lambda payload: payload["email"].update({"approved_drafts":3})), ("email-higher", lambda payload: payload["email"].update({"higher_approval_drafts":4})),
        ("email-total", lambda payload: payload["email"].update({"total_drafts":6})), ("email-sent", lambda payload: payload["email"].update({"sent_count":1,"state":"SENT"})),
        ("tracker-approved", lambda payload: payload["tracker_updates"].update({"approved_status_drafts":3})), ("tracker-higher", lambda payload: payload["tracker_updates"].update({"higher_approval_status_drafts":4})),
        ("tracker-total", lambda payload: payload["tracker_updates"].update({"total_drafts":6})), ("tracker-applied", lambda payload: payload["tracker_updates"].update({"applied_count":1,"state":"APPLIED"})),
        ("source-modified", lambda payload: payload["tracker_updates"].update({"source_modified":True})), ("silent-hold", lambda payload: payload["no_action"].update({"silent_hold":[]})),
        ("silent-status", lambda payload: payload["no_action"].update({"source_status":"Approved"})), ("silent-email", lambda payload: payload["no_action"].update({"email_draft_count":1})),
        ("silent-update", lambda payload: payload["no_action"].update({"tracker_update_draft_count":1})), ("silent-slack", lambda payload: payload["no_action"].update({"slack_draft_count":1})),
        ("slack-draft", lambda payload: payload["slack"].update({"draft_count":1})), ("slack-post", lambda payload: payload["slack"].update({"posted_count":1,"state":"POSTED"})),
        ("trace", lambda payload: payload.update({"source_trace":[]})),
    ]
    for name, mutator in execution_mutations:
        criterion = "R010" if name.startswith("silent") else "R014"
        cases.append((f"execution-{name}", lambda root, m=mutator: mutate_json(root / EXECUTION, m), criterion))

    word_cases = [
        ("summary-approved","4","5","APPROVED","R016"),("summary-higher","3","2","HIGHER_APPROVAL","R016"),("summary-silent","1","0","SILENT_HOLD","R016"),
        ("alice-amount","$62.00","$162.00","Alice Johansson","R005"),("ben-status","Pending CFO Approval","Approved","Ben Torres","R006"),
        ("claudia-role","VP Travel","CFO","Claudia Reeves","R007"),("david-amount","$350.00","$550.00","David Osei","R005"),
        ("elena-receipt","No","Yes","Elena Marchetti","R010"),("elena-decision","SILENT_HOLD","APPROVED","Elena Marchetti","R010"),
        ("finn-role","IT Director","CFO","Finn Larsson","R008"),("grace-preapproval","YES","NO","Grace Pham","R009"),("hector-amount","$450.00","$4,500.00","Hector Kim","R005"),
        ("email-state","DRAFT_NOT_SENT","SENT","Emails","R014"),("tracker-state","DRAFT_NOT_APPLIED","APPLIED","Tracker updates","R014"),
        ("slack-state","NOT_AUTHORIZED_CHANNEL","POSTED","Slack","R014"),("side-effect","OFFLINE_ONLY","EXECUTED","Real side effects","R014"),
    ]
    for name, old, new, row_token, criterion in word_cases:
        cases.append((f"word-{name}", lambda root, o=old, n=new, token=row_token: replace_docx_text(root / DOCX, o, n, row_token=token), criterion))
    cases.extend([
        ("word-receipt-policy", lambda root: replace_docx_text(root / DOCX, "requires a receipt for every expense", "allows missing receipts below $100"), "R016"),
        ("word-precedence", lambda root: replace_docx_text(root / DOCX, "Missing Receipt -> Valid Pre-approval -> Category Rule -> Approve", "Approve -> Missing Receipt"), "R016"),
        ("word-elena-silent", lambda root: replace_docx_text(root / DOCX, "Elena Marchetti remains Pending with no notification or update", "Elena Marchetti is approved"), "R010"),
    ])

    ppt_cases = [
        ("summary-reviewed","8 reviewed","7 reviewed","R015"),("summary-approved","4 approved","5 approved","R015"),("summary-higher","3 higher approval","2 higher approval","R015"),
        ("summary-silent","1 silent hold","0 silent hold","R015"),("receipt-policy","every expense requires a receipt","receipts optional below $100","R015"),
        ("precedence","Missing Receipt -> Valid Pre-approval -> Category Rule -> Approve","Approve -> Missing Receipt","R015"),
        ("sent","no email sent","email sent","R014"),("applied","no tracker update applied","tracker update applied","R014"),
        ("slack","no Slack message","Slack message posted","R014"),("side-effect","real side effects=0","real side effects=1","R014"),
    ]
    for name, old, new, criterion in ppt_cases:
        cases.append((f"ppt-{name}", lambda root, o=old, n=new: replace_pptx_text(root / PPTX, o, n), criterion))
    cases.extend([
        ("ppt-alice-amount", lambda root: replace_pptx_row_text(root / PPTX, "Alice Johansson", "$62.00", "$162.00"), "R005"),
        ("ppt-ben-role", lambda root: replace_pptx_row_text(root / PPTX, "Ben Torres", "CFO", "VP Travel"), "R006"),
        ("ppt-claudia-decision", lambda root: replace_pptx_row_text(root / PPTX, "Claudia Reeves", "HIGHER_APPROVAL", "APPROVED"), "R007"),
        ("ppt-elena-decision", lambda root: replace_pptx_row_text(root / PPTX, "Elena Marchetti", "SILENT_HOLD", "APPROVED"), "R010"),
        ("ppt-finn-status", lambda root: replace_pptx_row_text(root / PPTX, "Finn Larsson", "Pending IT Director Approval", "Approved"), "R008"),
        ("ppt-grace-preapproval", lambda root: replace_pptx_row_text(root / PPTX, "Grace Pham", "YES", "NO"), "R009"),
        ("ppt-hector-amount", lambda root: replace_pptx_row_text(root / PPTX, "Hector Kim", "$450.00", "$4,500.00"), "R005"),
        ("ppt-drop-slide", lambda root: drop_ppt_slide(root / PPTX), "R015"),
    ])
    return cases


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--json-out", type=Path)
    args = parser.parse_args()
    judge = load_judge()
    cases = build_cases()
    results = []
    failures = []
    for name, mutator, expected_failure in cases:
        with tempfile.TemporaryDirectory(prefix=f"negative-5119-{name}-") as temporary:
            artifact_root = Path(temporary) / "artifacts"
            shutil.copytree(GOLDEN, artifact_root)
            try:
                mutator(artifact_root)
                result = judge.evaluate(artifact_root)
                detected = not result.get("criteria", {}).get(expected_failure, {}).get("passed", True)
                rejected = not result.get("passed", False) and float(result.get("reward", 1.0)) < 0.8
                results.append({"case":name,"expected_failure":expected_failure,"detected":detected,"rejected":rejected,"reward":result.get("reward"),"hurdle_failures":result.get("hurdle_failures",[])})
                if not detected or not rejected:
                    failures.append(f"{name}: expected={expected_failure}, detected={detected}, reward={result.get('reward')}")
            except Exception as exc:
                failures.append(f"{name}: mutation/evaluation error: {type(exc).__name__}: {exc}")
    payload = {
        "task_id":"cowork-evo/automationbench-5119-expense-report-control-package","case_count":len(results),"declared_case_count":len(cases),
        "all_rejected":not failures and len(results) == len(cases),"max_reward":max((float(item["reward"]) for item in results), default=None),"failures":failures,"results":results,
    }
    if args.json_out:
        args.json_out.parent.mkdir(parents=True, exist_ok=True)
        args.json_out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if failures or len(results) != len(cases):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
