#!/usr/bin/env python3
import json,mailbox,shutil,subprocess,tempfile
from email.generator import BytesGenerator
from email.policy import SMTP
from pathlib import Path
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
ROOT=Path(__file__).resolve().parents[1];G=ROOT/"solution/Golden";PY="REMOVED_LOCAL_PATH"
def score(c):
 r=c.parent/(c.name+".json");w=c.parent/(c.name+".txt");subprocess.run([PY,str(ROOT/"tests/judge.py"),"--artifact-dir",str(c),"--result",str(r),"--reward-file",str(w)],check=True,capture_output=True,text=True,timeout=240);return float(w.read_text())
def rewrite(p,mut):
 b=mailbox.mbox(p,create=False);ms=list(b);b.close();mut(ms)
 with p.open("wb") as h:
  for i,m in enumerate(ms):h.write(f"From x Sun Mar 15 09:00:{i:02d} 2026\n".encode());BytesGenerator(h,policy=SMTP,mangle_from_=True,maxheaderlen=0).flatten(m);h.write(b"\n")
def body(m,t):
 for h in ("Content-Type","Content-Transfer-Encoding","MIME-Version"):
  if h in m:del m[h]
 m.set_payload(t)
def main():
 with tempfile.TemporaryDirectory(prefix="salary5011-neg-") as tmp:
  r=Path(tmp);names=["missing","wrong_count","alice_flagged","carol_omitted","rounded","wrong_recipient","wrong_flag","applied","source_updated","formula_removed","word_false","ppt_omit"];cs=[]
  for n in names:c=r/n;shutil.copytree(G,c);cs.append(c)
  (cs[0]/"Salary_Band_Compliance_Review.pptx").unlink()
  p=json.loads((cs[1]/"Salary_Band_Execution_Control.json").read_text());p["flagged_employees"]=2;(cs[1]/"Salary_Band_Execution_Control.json").write_text(json.dumps(p))
  rewrite(cs[2]/"Compensation_Team_Report_Draft.mbox",lambda ms:body(ms[0],"4 employees: Alice Park, Bob Chen, Carol Diaz, Eve Liu. DRAFT_NOT_SENT."))
  rewrite(cs[3]/"Compensation_Team_Report_Draft.mbox",lambda ms:body(ms[0],"3 employees: Bob Chen 125000; Eve Liu 215000. DRAFT_NOT_SENT."))
  rewrite(cs[4]/"Compensation_Team_Report_Draft.mbox",lambda ms:body(ms[0],"3 employees: Bob Chen 125k; Carol Diaz 108k; Eve Liu 215k. DRAFT_NOT_SENT."))
  rewrite(cs[5]/"Compensation_Team_Report_Draft.mbox",lambda ms:ms[0].replace_header("To","wrong@company.example.com"))
  p=json.loads((cs[6]/"Salary_Flag_Update_Drafts.json").read_text());p["updates"][0]["flag"]="Under";(cs[6]/"Salary_Flag_Update_Drafts.json").write_text(json.dumps(p))
  p=json.loads((cs[7]/"Salary_Flag_Update_Drafts.json").read_text());p["applied_count"]=3;(cs[7]/"Salary_Flag_Update_Drafts.json").write_text(json.dumps(p))
  p=json.loads((cs[8]/"Salary_Band_Execution_Control.json").read_text());p["source_rows_updated"]=3;(cs[8]/"Salary_Band_Execution_Control.json").write_text(json.dumps(p))
  w=load_workbook(cs[9]/"Salary_Band_Compliance_Audit.xlsx");w["Band_Check"]["J2"]="Within";w.save(cs[9]/"Salary_Band_Compliance_Audit.xlsx")
  d=Document(cs[10]/"Salary_Band_Compliance_Brief.docx");d.add_paragraph("Alice Park is Over band.");d.save(cs[10]/"Salary_Band_Compliance_Brief.docx")
  p=Presentation(cs[11]/"Salary_Band_Compliance_Review.pptx")
  for s in p.slides:
   for sh in s.shapes:
    if getattr(sh,"has_text_frame",False):sh.text=sh.text.replace("Carol Diaz","Unknown")
    if getattr(sh,"has_table",False):
     for row in sh.table.rows:
      for cell in row.cells:cell.text=cell.text.replace("Carol Diaz","Unknown")
  p.save(cs[11]/"Salary_Band_Compliance_Review.pptx")
  x={c.name:score(c) for c in cs};print(json.dumps(x,indent=2));assert all(v==0 for v in x.values()),x
if __name__=="__main__":main()
