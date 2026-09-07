from __future__ import annotations

import argparse
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

SOURCES = [
    ("INTAKE", "deal-intake-sheet.docx", "当事方与生效日唯一来源"),
    ("CSA", "cloud-service-agreement.docx", "主协议、Order Form 与 Key Terms"),
    ("ORDER", "order-form.docx", "已选项与未填商务变量"),
    ("DPA", "dpa-module.docx", "数据保护义务与优先级"),
]

TERMS = [
 ("T01","当事方","Provider = Acme Cloud Services, Inc.; Customer = Wayne Enterprises, Inc.","INTAKE","已明确","不得倒置角色"),
 ("T02","生效日","June 22, 2026","INTAKE","已明确","各 Cover Page 应一致"),
 ("T03","Order Form 优先级","Order Form 对该 Agreement 的不一致之处优先于 Framework Terms","CSA Order Form 行1","已明确","不能反向表述"),
 ("T04","DPA 优先级","EEA SCCs/UK Addendum > DPA > Agreement","DPA §7","已明确","三层顺序必须完整"),
 ("T05","处理者角色","Customer 为 Controller 时 Provider 为 Processor；Customer 为 Processor 时 Provider 为 Subprocessor","DPA §1","已明确","保留两种关系"),
 ("T06","Subprocessor 变更","至少提前 10 个工作日通知；Customer 在通知后 30 日内可异议","DPA §2.4","已明确","10 business days 与 30 days 不得混淆"),
 ("T07","Security Incident","Provider 应无不当延迟通知，且最迟不超过知悉后 72 小时","DPA §4","已明确","不是 72 个工作小时"),
 ("T08","审计与记录","DPA 合规记录在 DPA 结束后保留 3 年；安全尽调请求每年最多 1 次","DPA §5.1/§5.3","已明确","分开两个时间约束"),
 ("T09","主协议删除","Customer 请求时，Provider 在终止后 60 日内删除 Customer Content","CSA §5.5(b)","已明确","与 DPA 到期机制并列"),
 ("T10","DPA 到期处置","DPA 到期后按 Customer 指示返还或删除；法律要求保留时例外","DPA §6.2","已明确","不得简化为无条件 60 日"),
 ("T11","Payment Dispute","双方在 15 日内共同解决；无争议金额仍按时支付","CSA §4.6","已明确","保留支付义务"),
 ("T12","重大违约补救","经 30 日通知后仍未补救，守约方可终止","CSA §5.3(a)","已明确","30 日是 cure notice 期"),
 ("T13","破产程序","破产/接管程序持续超过 60 日可通知终止","CSA §5.3(b)","已明确","区别于 30 日补救期"),
 ("T14","Force Majeure","使 Product 重大无法运行连续 30 日或以上，可终止受影响 Order Form 并按比例退预付费","CSA §5.4","已明确","保留退费后果"),
 ("T15","Warranty Remedy","Customer 发现后 45 日内通知；Provider 收到充分详情后 45 日内尝试恢复","CSA §6.4","已明确","保留两个 45 日节点"),
 ("T16","Liability Caps","General Cap 与 Increased Cap 已勾选计算方式，但倍数仍为 [Fill in a number]","CSA Key Terms 行5/7","未明确","不得臆造倍数或金额"),
 ("T17","Governing Law / Courts","州、国家与法院所在地均未填","CSA Key Terms 行2/3","未明确","不得借用其他模板的 Delaware"),
 ("T18","Order economics","Cloud Service、Subscription Period、Fees、自动扣款频率、非续期通知天数均未填","ORDER 行2/6/7/8/9","未明确","已勾 automatic payment 与 auto-renewal"),
 ("T19","DPA Cover Page","Agreement、subprocessor list、security contact、联系信息、数据类别与处理详情均未填","DPA Cover Page 行1–27","未明确","转移附录不完整"),
 ("T20","签署与通知","Order Form 与 DPA 签名、姓名、职务、日期及 Notice Address 均未填","ORDER/DPA 签署表","未明确","不得从 intake 推断签署权"),
]

def shade(cell,color):
 p=cell._tc.get_or_add_tcPr(); e=p.find(qn("w:shd")) or OxmlElement("w:shd"); e.set(qn("w:fill"),color)
 if e.getparent() is None:p.append(e)
def header(row):
 p=row._tr.get_or_add_trPr(); e=OxmlElement("w:tblHeader"); e.set(qn("w:val"),"true"); p.append(e)
def nosplit(row): row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
def cell(c,text,bold=False,size=8):
 c.text=""; r=c.paragraphs[0].add_run(text); r.bold=bold; r.font.size=Pt(size); c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
def bookmark(p,name,bid):
 s=OxmlElement("w:bookmarkStart"); s.set(qn("w:id"),str(bid)); s.set(qn("w:name"),name); e=OxmlElement("w:bookmarkEnd"); e.set(qn("w:id"),str(bid)); p._p.insert(0,s); p._p.append(e)
def link(p,text,anchor):
 h=OxmlElement("w:hyperlink"); h.set(qn("w:anchor"),anchor); r=OxmlElement("w:r"); rp=OxmlElement("w:rPr"); co=OxmlElement("w:color"); co.set(qn("w:val"),"0563C1"); u=OxmlElement("w:u"); u.set(qn("w:val"),"single"); rp.extend([co,u]); r.append(rp); t=OxmlElement("w:t"); t.text=text; r.append(t); h.append(r); p._p.append(h)
def field(p,code):
 f=OxmlElement("w:fldSimple"); f.set(qn("w:instr"),code); r=OxmlElement("w:r"); t=OxmlElement("w:t"); t.text="1"; r.append(t); f.append(r); p._p.append(f)

def verify(inp):
 d=Document(inp/"deal-intake-sheet.docx"); v={r.cells[0].text.strip():r.cells[1].text.strip() for r in d.tables[0].rows[1:]}; assert v=={"Customer Name":"Wayne Enterprises, Inc.","Provider Name":"Acme Cloud Services, Inc.","Effective Date":"June 22, 2026"}
 c="\n".join(p.text for p in Document(inp/"cloud-service-agreement.docx").paragraphs); a="\n".join(p.text for p in Document(inp/"dpa-module.docx").paragraphs)
 assert all(x in c for x in ["within 15 days","within 60 days","within 45 days"]); assert all(x in a for x in ["10 business days","30 days after notice","no later than 72 hours","for 3 years"])

def build(inp,outdir):
 verify(inp); d=Document(); s=d.sections[0]; s.top_margin=s.bottom_margin=Inches(.65); s.left_margin=s.right_margin=Inches(.68)
 d.styles["Normal"].font.name="Aptos"; d.styles["Normal"].font.size=Pt(9); d.styles["Title"].font.color.rgb=RGBColor(31,78,121)
 p=d.add_paragraph(style="Title"); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run("Cloud Agreement Terms & Readiness Audit")
 p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.add_run("Acme Cloud Services, Inc. / Wayne Enterprises, Inc.\nEffective Date: June 22, 2026").bold=True
 p=d.add_paragraph("Evidence-bound summary of four frozen inputs. Not legal advice. Unfilled variables remain unresolved."); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
 d.add_heading("Navigation",1); field(d.add_paragraph(),'TOC \\o "1-2" \\h \\z \\u')
 d.add_heading("1. Executive readiness snapshot",1); t=d.add_table(rows=5,cols=2); t.style="Table Grid"
 snap=[("Provider","Acme Cloud Services, Inc."),("Customer","Wayne Enterprises, Inc."),("Effective Date","June 22, 2026"),("Controls reviewed","20"),("Readiness","15 evidenced controls; 5 unresolved control groups; not signature-ready")]
 for r,v in zip(t.rows,snap): cell(r.cells[0],v[0],True,9); shade(r.cells[0],"D9EAF7"); cell(r.cells[1],v[1],False,9); nosplit(r)
 d.add_heading("2. Source register and reading rule",1); t=d.add_table(rows=1,cols=3); t.style="Table Grid"
 for c,h in zip(t.rows[0].cells,["Code","Frozen file","Use"]): cell(c,h,True); shade(c,"1F4E79")
 header(t.rows[0])
 for v in SOURCES:
  r=t.add_row(); nosplit(r)
  for c,x in zip(r.cells,v):cell(c,x)
 d.add_paragraph("A precedence clause controls only within its stated scope. No blank variable is inferred from drafting examples or another public template.")
 d.add_section(WD_ORIENT.LANDSCAPE); s=d.sections[-1]; s.orientation=WD_ORIENT.LANDSCAPE; s.page_width,s.page_height=Inches(11.69),Inches(8.27); s.top_margin=s.bottom_margin=Inches(.55); s.left_margin=s.right_margin=Inches(.45)
 d.add_heading("3. Terms, parties and readiness matrix",1); m=d.add_table(rows=1,cols=6); m.style="Table Grid"; hs=["ID","Topic","Evidence-bound term","Source locator","Status","Implication"]
 for c,h in zip(m.rows[0].cells,hs):cell(c,h,True,7);shade(c,"1F4E79")
 header(m.rows[0])
 for i,v in enumerate(TERMS,1):
  r=m.add_row();nosplit(r)
  for c,x in zip(r.cells,v):cell(c,x,False,7)
  shade(r.cells[4],"E2F0D9" if v[4]=="已明确" else "FCE4D6");r.cells[0].text="";link(r.cells[0].paragraphs[0],v[0],f"evidence_{v[0]}")
 d.add_heading("4. Signature-readiness actions",1); acts=[("A1","Legal","Complete governing law, courts and both cap multipliers.","T16-T17"),("A2","Commercial","Complete service, term, fees, billing cadence and non-renewal days.","T18"),("A3","Privacy/Security","Complete DPA contacts, transfers, data categories and processing details.","T19"),("A4","Authorized signers","Complete signatures and notice addresses; do not infer authority.","T20"),("A5","Deal Operations","Propagate exact parties/date after A1-A4, then recheck conflicts.","T01-T04")]
 a=d.add_table(rows=1,cols=4);a.style="Table Grid"
 for c,h in zip(a.rows[0].cells,["Action","Owner","Required resolution","Controls"]):cell(c,h,True);shade(c,"1F4E79")
 header(a.rows[0])
 for v in acts:
  r=a.add_row();nosplit(r)
  for c,x in zip(r.cells,v):cell(c,x)
 d.add_heading("5. Evidence excerpts",1)
 for i,v in enumerate(TERMS,1):
  p=d.add_paragraph(style="List Number");p.add_run(f"{v[0]} | {v[3]}: ").bold=True;p.add_run(v[2]);bookmark(p,f"evidence_{v[0]}",i)
 for sec in d.sections:
  sec.header.is_linked_to_previous=False;sec.footer.is_linked_to_previous=False;hp=sec.header.paragraphs[0];hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT;hp.text="Agreement Terms & Readiness Audit | 2026-06-22";fp=sec.footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.CENTER;fp.add_run("Page ");field(fp,"PAGE");fp.add_run(" of ");field(fp,"NUMPAGES")
 outdir.mkdir(parents=True,exist_ok=True);o=outdir/"agreement_terms_audit.docx";d.save(o)

def main():
 p=argparse.ArgumentParser();p.add_argument("--input-dir",required=True);p.add_argument("--output-dir",required=True);a=p.parse_args();build(Path(a.input_dir),Path(a.output_dir))
if __name__=="__main__":main()
