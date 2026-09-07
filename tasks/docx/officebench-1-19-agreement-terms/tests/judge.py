from __future__ import annotations
import argparse,hashlib,json,re,zipfile
from pathlib import Path
from docx import Document
from lxml import etree

W={"R001":8,"R002":10,"R003":12,"R004":22,"R005":16,"R006":8,"R007":8,"R008":8,"R009":5,"R010":3}
H={"R001":0,"R002":.35,"R003":.55,"R004":.65,"R005":.70,"R009":.75}
NS={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
def item(cid,score,evidence):return {"criterion_id":cid,"score":round(score,6),"evidence":evidence,"method":"deterministic","raw":{"weight":W[cid]}}
def norm(s):return re.sub(r"\s+"," ",s).strip()
def source_truth(root):
 d=Document(root/"deal-intake-sheet.docx"); vals={r.cells[0].text.strip():r.cells[1].text.strip() for r in d.tables[0].rows[1:]}
 c=norm(" ".join(p.text for p in Document(root/"cloud-service-agreement.docx").paragraphs)); a=norm(" ".join(p.text for p in Document(root/"dpa-module.docx").paragraphs)); o=norm(" ".join(c.text for t in Document(root/"order-form.docx").tables for r in t.rows for c in r.cells))
 assert vals=={"Customer Name":"Wayne Enterprises, Inc.","Provider Name":"Acme Cloud Services, Inc.","Effective Date":"June 22, 2026"}
 proofs=["within 15 days","within 60 days","following 30 days notice","continue for more than 60 days","30 or more consecutive days","within 45 days"]
 assert all(x in c for x in proofs) and all(x in a for x in ["10 business days","30 days after notice","no later than 72 hours","for 3 years","once a year"])
 assert "Automatic payment" in o and "Non-Renewal Notice Date" in o
 return vals
def check(root,outdir):
 files=sorted(p.name for p in outdir.iterdir()) if outdir.is_dir() else []; out=outdir/"agreement_terms_audit.docx"; rs=[]
 valid=files==[out.name] and out.is_file() and zipfile.is_zipfile(out)
 rs.append(item("R001",1 if valid else 0,f"files={files}; native_docx={valid}"))
 if not valid:return rs
 vals=source_truth(root); d=Document(out); text=norm("\n".join(p.text for p in d.paragraphs)+"\n"+"\n".join(c.text for t in d.tables for r in t.rows for c in r.cells))
 party=all(x in text for x in vals.values()) and "Provider = Acme Cloud Services, Inc.; Customer = Wayne Enterprises, Inc." in text
 rs.append(item("R002",1 if party else 0,f"party/date exact={party}"))
 precedence=["Order Form 对该 Agreement 的不一致之处优先于 Framework Terms","EEA SCCs/UK Addendum > DPA > Agreement","Customer 为 Controller 时 Provider 为 Processor","Customer 为 Processor 时 Provider 为 Subprocessor"]
 pc=sum(x in text for x in precedence);rs.append(item("R003",pc/4,f"precedence/roles={pc}/4"))
 matrix=next((t for t in d.tables if len(t.rows)==21 and len(t.columns)==6),None); ids=[] if matrix is None else [r.cells[0].text.strip() for r in matrix.rows[1:]]; exact=ids==[f"T{i:02}" for i in range(1,21)]
 statuses=[] if matrix is None else [r.cells[4].text.strip() for r in matrix.rows[1:]]; stat=statuses==["已明确"]*15+["未明确"]*5
 rs.append(item("R004",(int(exact)+int(stat))/2,f"matrix={getattr(matrix,'_tbl',None) is not None}; ids={exact}; status={stat}"))
 facts=["10 个工作日","30 日内可异议","72 小时","3 年","60 日内删除","按 Customer 指示返还或删除","法律要求保留时例外","15 日内","30 日通知","60 日可通知","连续 30 日","45 日内通知","45 日内尝试恢复","[Fill in a number]","州、国家与法院所在地均未填"]
 fc=sum(x in text for x in facts);rs.append(item("R005",fc/len(facts),f"critical facts={fc}/{len(facts)}"))
 actions=next((t for t in d.tables if len(t.rows)==6 and len(t.columns)==4),None); ac=0 if actions is None else sum(f"A{i}" in actions.rows[i].cells[0].text for i in range(1,6));rs.append(item("R006",ac/5,f"actions={ac}/5"))
 with zipfile.ZipFile(out) as z:
  xml=etree.fromstring(z.read("word/document.xml")); names=set(xml.xpath("//w:bookmarkStart/@w:name",namespaces=NS)); anchors=xml.xpath("//w:hyperlink/@w:anchor",namespaces=NS); sect=xml.xpath("//w:sectPr",namespaces=NS); landscape=xml.xpath("//w:sectPr/w:pgSz[@w:orient='landscape']",namespaces=NS); cant=xml.xpath("//w:cantSplit",namespaces=NS)
  fields=" ".join(" ".join(etree.fromstring(z.read(n)).xpath("//w:fldSimple/@w:instr | //w:instrText/text()",namespaces=NS)) for n in z.namelist() if n.startswith("word/") and n.endswith(".xml"))
 nav=sum(f"evidence_T{i:02}" in names and f"evidence_T{i:02}" in anchors for i in range(1,21));rs.append(item("R007",nav/20,f"bookmark-links={nav}/20"))
 structural=sum([len(d.tables)==4,len(sect)>=2,bool(landscape),"TOC" in fields,"PAGE" in fields and "NUMPAGES" in fields,len(cant)>=30]);rs.append(item("R008",structural/6,f"tables={len(d.tables)}; sections={len(sect)}; landscape={bool(landscape)}; fields={fields}; cantSplit={len(cant)}"))
 forbidden=["2x the Fees","3x the Fees","$1,000,000","laws of Delaware","Bruce Wayne","Jane Doe"]
 bad=[x for x in forbidden if x in text]; unresolved=all(x in text for x in ["不得臆造倍数或金额","不得借用其他模板的 Delaware","不得从 intake 推断签署权"]);rs.append(item("R009",1 if not bad and unresolved else 0,f"fabricated={bad}; unresolved safeguards={unresolved}"))
 headings=sum(p.style.name.startswith("Heading") for p in d.paragraphs);rs.append(item("R010",1 if headings>=6 and len(text)>2500 else .5,f"headings={headings}; chars={len(text)}"));return rs
def main():
 p=argparse.ArgumentParser();p.add_argument("--artifact-dir",type=Path,required=True);p.add_argument("--result",type=Path,required=True);p.add_argument("--reward-file",type=Path,required=True);a=p.parse_args();root=Path(__file__).parents[1]/"data/input_files"
 try:rs=check(root,a.artifact_dir)
 except Exception as e:rs=[item("R001",0,f"exception={type(e).__name__}: {e}")]
 scores={x["criterion_id"]:x["score"] for x in rs};reward=sum(W[k]*scores.get(k,0) for k in W)/100;fails=[k for k,c in H.items() if scores.get(k,0)<.999999];
 if fails:reward=min(reward,min(H[k] for k in fails))
 payload={"task_id":"cowork-evo/officebench-1-19-agreement-terms","rubric_version":"officebench-1-19-agreement-terms-v0.3","judge_version":"officebench-1-19-agreement-terms-deterministic-v0.3","artifact_sha256":[hashlib.sha256((a.artifact_dir/"agreement_terms_audit.docx").read_bytes()).hexdigest()] if (a.artifact_dir/"agreement_terms_audit.docx").is_file() else [],"criterion_results":rs,"reward":round(reward,6),"passed":reward>=.8 and not fails,"hurdle_failures":fails};a.result.parent.mkdir(parents=True,exist_ok=True);a.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2));a.reward_file.parent.mkdir(parents=True,exist_ok=True);a.reward_file.write_text(f"{reward:.6f}\n")
if __name__=="__main__":main()
