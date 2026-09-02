#!/usr/bin/env python3
"""5070 专属负例：覆盖范围、身份、PTO、状态、公式、叙事和执行边界。"""

import json
import mailbox
import shutil
import subprocess
import tempfile
from email.generator import BytesGenerator
from email.policy import SMTP
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


ROOT = Path(__file__).resolve().parents[1]
GOLDEN = ROOT / "solution/Golden"
PYTHON = "REMOVED_LOCAL_PATH"


def score(case):
    result = case.parent / f"{case.name}.json"
    reward = case.parent / f"{case.name}.txt"
    subprocess.run(
        [PYTHON, str(ROOT / "tests/judge.py"), "--artifact-dir", str(case), "--result", str(result), "--reward-file", str(reward)],
        check=True,
        capture_output=True,
        text=True,
        timeout=240,
    )
    return float(reward.read_text())


def rewrite_mbox(path, old, new):
    box = mailbox.mbox(path, create=False)
    messages = list(box)
    box.close()
    message = messages[0]
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                part.set_payload(part.get_payload().replace(old, new))
    else:
        message.set_payload(message.get_payload().replace(old, new))
    with path.open("wb") as handle:
        for index, item in enumerate(messages):
            handle.write(f"From x Wed Mar 18 09:00:{index:02d} 2026\n".encode())
            BytesGenerator(handle, policy=SMTP, mangle_from_=True, maxheaderlen=0).flatten(item)
            handle.write(b"\n")


def replace_docx(path, old, new):
    document = Document(path)
    for paragraph in document.paragraphs:
        if old in paragraph.text:
            for run in paragraph.runs:
                run.text = run.text.replace(old, new)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.text = run.text.replace(old, new)
    document.save(path)


def replace_pptx(path, old, new):
    presentation = Presentation(path)
    for slide in presentation.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.text = run.text.replace(old, new)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        cell.text = cell.text.replace(old, new)
    presentation.save(path)


def main():
    names = [
        "missing_file", "wrong_scope_count", "nadia_bamboo", "wrong_effective_date", "nadia_jira",
        "wrong_pto", "omit_ian_payroll", "wrong_tracker_status", "executed_action", "formula_removed",
        "wrong_reconciliation", "word_scope_false", "ppt_omit_nadia",
    ]
    with tempfile.TemporaryDirectory(prefix="offboarding5070-neg-") as temp:
        root = Path(temp)
        cases = []
        for name in names:
            case = root / name
            shutil.copytree(GOLDEN, case)
            cases.append(case)

        (cases[0] / "Offboarding_Operations_Review.pptx").unlink()

        payload = json.loads((cases[1] / "Offboarding_Execution_Control.json").read_text())
        payload["employees_in_scope"] = 3
        (cases[1] / "Offboarding_Execution_Control.json").write_text(json.dumps(payload))

        payload = json.loads((cases[2] / "BambooHR_Status_Update_Drafts.json").read_text())
        payload["updates"][1]["employee"] = "Nadia Petrova"
        payload["updates"][1]["employee_id"] = "EMP-2005"
        (cases[2] / "BambooHR_Status_Update_Drafts.json").write_text(json.dumps(payload))

        payload = json.loads((cases[3] / "BambooHR_Status_Update_Drafts.json").read_text())
        payload["updates"][0]["effective_date"] = "2026-03-22"
        (cases[3] / "BambooHR_Status_Update_Drafts.json").write_text(json.dumps(payload))

        payload = json.loads((cases[4] / "Jira_Access_Revocation_Drafts.json").read_text())
        payload["tickets"][1]["summary"] = "Revoke access for Nadia Petrova (EMP-2005)"
        (cases[4] / "Jira_Access_Revocation_Drafts.json").write_text(json.dumps(payload))

        payroll_path = cases[5] / "Payroll_Notification_Draft.mbox"
        payroll_path.write_bytes(payroll_path.read_bytes().replace(b"5 da=\nys", b"6 da=\nys"))
        rewrite_mbox(cases[6] / "Payroll_Notification_Draft.mbox", "Ian Costa", "Unknown Employee")

        payload = json.loads((cases[7] / "Tracker_Status_Update_Drafts.json").read_text())
        payload["updates"][0]["target_status"] = "Pending"
        (cases[7] / "Tracker_Status_Update_Drafts.json").write_text(json.dumps(payload))

        payload = json.loads((cases[8] / "Offboarding_Execution_Control.json").read_text())
        payload["external_actions_executed"] = 7
        payload["real_side_effects"] = 7
        (cases[8] / "Offboarding_Execution_Control.json").write_text(json.dumps(payload))

        workbook = load_workbook(cases[9] / "Offboarding_Multi_System_Audit.xlsx")
        workbook["Scope_Decisions"]["J2"] = "PROCESS_OFFBOARDING"
        workbook.save(cases[9] / "Offboarding_Multi_System_Audit.xlsx")

        workbook = load_workbook(cases[10] / "Offboarding_Multi_System_Audit.xlsx")
        workbook["Cross_System_Reconciliation"]["G2"] = "NO"
        workbook.save(cases[10] / "Offboarding_Multi_System_Audit.xlsx")

        replace_docx(cases[11] / "Offboarding_Operations_Brief.docx", "SKIP_OUT_OF_SCOPE", "PROCESS_OFFBOARDING")
        replace_pptx(cases[12] / "Offboarding_Operations_Review.pptx", "Nadia Petrova", "Unknown Employee")

        results = {case.name: score(case) for case in cases}
        print(json.dumps(results, indent=2))
        assert all(value == 0 for value in results.values()), results


if __name__ == "__main__":
    main()
