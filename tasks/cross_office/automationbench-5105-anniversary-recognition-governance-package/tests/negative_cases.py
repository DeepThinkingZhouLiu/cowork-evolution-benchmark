#!/usr/bin/env python3
import json,mailbox,shutil,subprocess,tempfile
from email.generator import BytesGenerator
from email.policy import SMTP
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
ROOT=Path(__file__).resolve().parents[1];GOLDEN=ROOT/"solution/Golden";PYTHON="REMOVED_LOCAL_PATH"
def score(case):
 result=case.parent/(case.name+".json");reward=case.parent/(case.name+".txt");subprocess.run([PYTHON,str(ROOT/"tests/judge.py"),"--artifact-dir",str(case),"--result",str(result),"--reward-file",str(reward)],check=True,capture_output=True,text=True,timeout=240);return float(reward.read_text())
def rewrite(path,mutation):
 box=mailbox.mbox(path,create=False);messages=list(box);box.close();mutation(messages)
 with path.open("wb") as handle:
  for index,message in enumerate(messages):handle.write(f"From x Wed Apr 15 09:00:{index:02d} 2026\n".encode());BytesGenerator(handle,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(message);handle.write(b"\n")
def main():
 with tempfile.TemporaryDirectory(prefix="anniversary5105-neg-") as temp:
  root=Path(temp);names=["missing","wrong_count","frank_march","frank_years","inactive_included","executive_cc","email_sent","wrong_to","formula_removed","word_false","ppt_omit","extra_file"];cases=[]
  for name in names:case=root/name;shutil.copytree(GOLDEN,case);cases.append(case)
  (cases[0]/"Anniversary_Review.pptx").unlink()
  payload=json.loads((cases[1]/"Anniversary_Execution_Control.json").read_text());payload["april_celebrants"]=4;(cases[1]/"Anniversary_Execution_Control.json").write_text(json.dumps(payload))
  payload=json.loads((cases[2]/"Hire_Date_Correction_Register.json").read_text());payload["corrected_hire_date"]="2024-03-01";(cases[2]/"Hire_Date_Correction_Register.json").write_text(json.dumps(payload))
  payload=json.loads((cases[3]/"Hire_Date_Correction_Register.json").read_text());payload["anniversary_years"]=3;(cases[3]/"Hire_Date_Correction_Register.json").write_text(json.dumps(payload))
  rewrite(cases[4]/"April_Anniversary_AllStaff_Draft.mbox",lambda ms:ms[0].set_payload(ms[0].get_payload()+"\nIrene Kowalski"))
  rewrite(cases[5]/"April_Anniversary_AllStaff_Draft.mbox",lambda ms:ms[0].__setitem__("Cc","executives@company.example.com"))
  payload=json.loads((cases[6]/"Anniversary_Execution_Control.json").read_text());payload["emails_sent"]=1;(cases[6]/"Anniversary_Execution_Control.json").write_text(json.dumps(payload))
  rewrite(cases[7]/"April_Anniversary_AllStaff_Draft.mbox",lambda ms:ms[0].replace_header("To","executives@company.example.com"))
  workbook=load_workbook(cases[8]/"Anniversary_Audit.xlsx");workbook["Anniversary_Decisions"]["M2"]="INCLUDE_APRIL_ANNIVERSARY";workbook.save(cases[8]/"Anniversary_Audit.xlsx")
  document=Document(cases[9]/"Anniversary_Brief.docx");document.add_paragraph("Executives were copied and the all-staff email was sent.");document.save(cases[9]/"Anniversary_Brief.docx")
  presentation=Presentation(cases[10]/"Anniversary_Review.pptx")
  for slide in presentation.slides:
   for shape in slide.shapes:
    if getattr(shape,"has_text_frame",False):shape.text=shape.text.replace("Frank Torres","Unknown")
    if getattr(shape,"has_table",False):
     for row in shape.table.rows:
      for cell in row.cells:cell.text=cell.text.replace("Frank Torres","Unknown")
  presentation.save(cases[10]/"Anniversary_Review.pptx");(cases[11]/"unrequested.txt").write_text("extra")
  rewards={case.name:score(case) for case in cases};print(json.dumps(rewards,indent=2));assert all(value==0 for value in rewards.values()),rewards
if __name__=="__main__":main()
