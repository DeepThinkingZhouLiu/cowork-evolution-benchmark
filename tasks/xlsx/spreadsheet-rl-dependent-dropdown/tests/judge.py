#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, re, shutil, subprocess, tempfile, zipfile
from datetime import timedelta
from pathlib import Path
from openpyxl import load_workbook

TASK_ID = "cowork-evo/spreadsheet-rl-dependent-dropdown"
KIND = 'dropdown'
INPUT = 'account_form_input.xlsx'
OUTPUT = 'account_form_completed.xlsx'
WEIGHTS = {"R001": .10, "R002": .35, "R003": .15, "R004": .15, "R005": .15, "R006": .05, "R007": .10, "R008": -.15}

def c(cid, score, evidence): return {"criterion_id": cid, "score": round(max(0.0, min(1.0, float(score))), 6), "evidence": evidence, "method": "deterministic"}
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def eq(a,b):
    if isinstance(a, (int,float)) and isinstance(b, (int,float)): return abs(float(a)-float(b)) <= 1e-6 * max(1.0, abs(float(a)), abs(float(b)))
    if isinstance(a, timedelta) and isinstance(b, timedelta): return abs((a-b).total_seconds()) < 1e-6
    return a == b or (a in (None, "") and b in (None, ""))
def numeric(value):
    if value is None or value == "": return None
    if isinstance(value, timedelta): return value.total_seconds()/3600.0
    if isinstance(value, (int,float)): return float(value)*24 if 0 <= float(value) < 2 else float(value)
    text=str(value).replace("\xa0"," ").strip().replace(",","")
    if text in ("", "-", "--"): return None
    if ":" in text:
        left,right=text.split(":",1)
        try: return float(left)+float(right)/60
        except ValueError:return None
    try:return float(text)
    except ValueError:return None
def calculated_workbook(path):
    """Load cached formula values after a headless LibreOffice round trip."""
    lo=shutil.which("libreoffice") or shutil.which("soffice")
    if lo is None: return load_workbook(path,data_only=True)
    with tempfile.TemporaryDirectory(prefix="xlsx-judge-") as td:
        work=Path(td); source_dir=work/"source"; recalc_dir=work/"recalc"; source_dir.mkdir(); recalc_dir.mkdir()
        source=source_dir/path.name; source.write_bytes(path.read_bytes())
        profile=(work/"profile").resolve().as_uri()
        subprocess.run([lo,"--headless",f"-env:UserInstallation={profile}","--convert-to","xlsx","--outdir",str(recalc_dir),str(source)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,check=False)
        converted=recalc_dir/path.name
        return load_workbook(converted if converted.exists() else source,data_only=True)
def source_preserved(src, out, allowed):
    bad=[]; checked=0
    for s in src.sheetnames:
        if s not in out.sheetnames: return 0.0, [f"missing sheet {s}"]
        a,b=src[s],out[s]
        for r in range(1,a.max_row+1):
            for col in range(1,a.max_column+1):
                key=f"{s}!{a.cell(r,col).coordinate}"
                if key in allowed: continue
                checked += 1
                if not eq(a.cell(r,col).value,b.cell(r,col).value): bad.append(key)
    return max(0.0, 1-len(bad)/max(checked,1)), bad[:20]
def dropdown_checks(out, src):
    if out.sheetnames != ["Account","Lists","Form"]: return (0,0,0,0,0, [f"sheets={out.sheetnames}"]) 
    account=src["Account"]; expected=[]; groups={}
    for r in range(2,account.max_row+1):
        cat,sub=account.cell(r,3).value,account.cell(r,4).value
        if cat not in groups: expected.append(cat); groups[cat]=[]
        if sub not in groups[cat]: groups[cat].append(sub)
    lists=out["Lists"]; values_ok=True; count=0
    for c,cat in enumerate(expected,1):
        if lists.cell(1,c).value != cat: values_ok=False
        for r,sub in enumerate(groups[cat],2): count+=1; values_ok &= lists.cell(r,c).value == sub
    names=out.defined_names; names_ok="Categories" in names and all(re.sub(r"[^A-Za-z0-9_]","_",str(cat).replace(" ","_")) in names for cat in expected)
    dvs={dv.sqref.__str__(): (dv.formula1 or "") for dv in out["Form"].data_validations.dataValidation}
    validation_ok=dvs.get("B2")=="Categories" and "INDIRECT" in dvs.get("C2","").upper()
    return float(values_ok),float(names_ok),float(validation_ok),float(len(expected)==7),float(count==sum(len(v) for v in groups.values())),[]
def reshape_expected(src):
    ws=src["DB"]; areas=[]; pairs=[]
    for col in range(2,ws.max_column+1):
        cat=re.sub(r"\d+$","",str(ws.cell(4,col).value or "")); area=ws.cell(5,col).value
        if area not in areas: areas.append(area)
        pairs.append((cat,area,col))
    rows=[]
    for cat in sorted(dict.fromkeys(x[0] for x in pairs)):
        for seq in range(1,6):
            row=[cat,seq]
            for area in areas:
                col=next((c for ccat,a,c in pairs if ccat==cat and a==area),None)
                row.append(numeric(ws.cell(seq+5,col).value) if col else None)
            rows.append(row)
    return ["Removed Characters","Attribute"]+areas,rows
def reshape_checks(out, src):
    if out.sheetnames != ["Sheet1"]: return 0,0,0,0,0, ["unexpected sheets"]
    headers, rows=reshape_expected(src); ws=out["Sheet1"]; got=[[ws.cell(r,c).value for c in range(1,len(headers)+1)] for r in range(2,2+len(rows))]
    matches=sum(eq(a,b) for arow,brow in zip(rows,got) for a,b in zip(arow,brow)); total=len(rows)*len(headers)
    return float(ws.max_row==21 and ws.max_column==10),matches/max(total,1),float(got==rows),float(len(ws.merged_cells.ranges)==0),float("#REF!" not in str(got)),[]
def hours_expected(src):
    ws=src["Data"]; months=[ws.cell(1,c).value for c in range(3,ws.max_column)]; rows=[]
    for r in range(3,25):
        vals=[numeric(ws.cell(r,c).value) or 0 for c in range(3,ws.max_column)]; rows.append([ws.cell(r,2).value]+vals+[sum(vals)])
    return ["Department"]+months+["TOTAL"],rows
def hours_checks(out, src):
    if "Analysis" not in out.sheetnames: return 0,0,0,0,0,["Analysis missing"]
    headers,rows=hours_expected(src); ws=out["Analysis"]
    got_wb=calculated_workbook(_OUT_PATH); got=got_wb["Analysis"]
    matches=0; total=len(rows)*len(headers)
    for r,row in enumerate(rows,2):
        for col,expected in enumerate(row,1):
            actual=got.cell(r,col).value; matches += int(eq(actual,expected))
    formula_count=sum(1 for r in range(2,24) for col in range(2,31) if isinstance(ws.cell(r,col).value,str) and ws.cell(r,col).value.startswith("="))
    chart_count=len(ws._charts); source_ok=out["Data"].max_row==src["Data"].max_row and out["Data"].max_column==src["Data"].max_column
    got_wb.close(); return float(ws.max_row>=25 and ws.max_column==30),matches/max(total,1),float(formula_count>=22*29),float(chart_count==1),float(source_ok),[]
def outer_expected(src):
    maps=[]; keys=[]
    for s in ("Table1","Table2"):
        ws=src[s]; m={ws.cell(r,2).value:[ws.cell(r,c).value for c in range(1,6)] for r in range(2,ws.max_row+1)}; maps.append(m)
        for key in m:
            if key not in keys: keys.append(key)
    rows=[]
    for key in keys:
        a,b=maps[0].get(key,[None]*5),maps[1].get(key,[None]*5); rows.append([key]+a+[None]*0+b+[a[0] or b[0],a[1] or b[1]])
    return keys,maps
def outer_checks(out, src):
    if out.sheetnames != ["Table1","Table2","Output"]: return 0,0,0,0,0,["unexpected sheets"]
    keys,maps=outer_expected(src); ws=out["Output"]; calc_wb=calculated_workbook(_OUT_PATH); calc_ws=calc_wb["Output"]; key_ok=[ws.cell(r,1).value for r in range(2,2+len(keys))]==keys; core=0; total=len(keys)*13; formulas=0
    for r,key in enumerate(keys,2):
        a=maps[0].get(key,[None]*5); b=maps[1].get(key,[None]*5); expected=[key]+a+b+[a[0] or b[0],a[1] or b[1]]
        for col,exp in enumerate(expected,1):
            actual=ws.cell(r,col).value
            if col==1: ok=eq(actual,exp)
            elif isinstance(actual,str) and actual.startswith("="): formulas+=1; ok=eq(calc_ws.cell(r,col).value,exp)
            else: ok=eq(actual,exp)
            core += int(ok)
    calc_wb.close(); pres,_=source_preserved(src,out,set()); return float(key_ok and core==total),core/max(total,1),float(formulas>=len(keys)*10),pres,float(len(ws._charts)==0),[]
def archer_expected(src):
    ws=src["NEW"]; groups=[(3,4),(6,7),(9,10),(12,13),(15,17),(19,20)]; rows=[]
    for r in range(3,33):
        maxima=[]
        for left,right in groups:
            vals=[ws.cell(r,c).value for c in range(left,right+1) if isinstance(ws.cell(r,c).value,(int,float))]
            maxima.append(max(vals) if vals else 0)
        top=sorted(enumerate(maxima), key=lambda x:(-x[1],x[0]))[:4]; values=[x[1] for x in top]
        positions=[]
        for idx,value in top:
            group=groups[idx]; matches=[c for c in range(group[0],group[1]+1) if ws.cell(r,c).value==value]; positions.append(matches[-1] if matches else None)
        while len(values)<4: values.append(0); positions.append(None)
        rows.append((values,positions))
    return rows
def archer_checks(out,src):
    if out.sheetnames != ["NEW"]: return 0,0,0,0,0,["unexpected sheets"]
    ws=out["NEW"]; source=src["NEW"]; expected=archer_expected(src); score=position=formulas=0; total=120
    cached_wb=calculated_workbook(_OUT_PATH); cached_ws=cached_wb["NEW"]
    for i,r in enumerate(range(3,33)):
        vals,pos=expected[i]
        for j,col in enumerate(range(25,29)):
            actual=ws.cell(r,col).value; cached=cached_ws.cell(r,col).value
            score += int(eq(cached,vals[j]) or eq(actual,vals[j]))
            formulas += int(isinstance(actual,str) and actual.startswith("=") and "LARGE" in actual.upper())
        for j,col in enumerate(range(30,34)):
            actual=ws.cell(r,col).value; cached=cached_ws.cell(r,col).value
            position += int(eq(cached,pos[j]) or eq(actual,pos[j]))
            formulas += int(isinstance(actual,str) and actual.startswith("="))
    cf_ok=any(str(key).find("C3:T32")>=0 for key in ws.conditional_formatting)
    pres,_=source_preserved(src,out,{f"NEW!{__import__('openpyxl').utils.get_column_letter(c)}{r}" for r in range(3,33) for c in range(25,34)})
    cached_wb.close(); return float(score>=110 and position>=110),float((score+position)/240),float(formulas>=240),float(cf_ok),pres,[]
def main():
    global _OUT_PATH
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir",type=Path,required=True); ap.add_argument("--result",type=Path,required=True); ap.add_argument("--reward-file",type=Path,required=True); args=ap.parse_args(); root=Path(__file__).parents[1]; src_path=root/"data/input_files"/INPUT; _OUT_PATH=args.artifact_dir/OUTPUT
    native=_OUT_PATH.is_file() and zipfile.is_zipfile(_OUT_PATH); results=[c("R001",native,"output exists and is native OOXML")]
    if not native:
        results += [c(k,0,"missing or corrupt output") for k in ("R002","R003","R004","R005","R006","R008")]; results.append(c("R007",0,"invalid output penalty"))
    else:
        try:
            src=load_workbook(src_path,data_only=False); out=load_workbook(_OUT_PATH,data_only=False)
            if KIND=="dropdown": a=dropdown_checks(out,src); criteria=(a[0],a[1],a[2],a[3],a[4],0,[],[])
            elif KIND=="reshape": a=reshape_checks(out,src); criteria=(a[0],a[1],a[2],a[3],a[4],0,[],[])
            elif KIND=="hours": a=hours_checks(out,src); criteria=(a[0],a[1],a[2],a[3],a[4],0,[],[])
            elif KIND=="outer": a=outer_checks(out,src); criteria=(a[0],a[1],a[2],a[3],a[4],0,[],[])
            else: a=archer_checks(out,src); criteria=(a[0],a[1],a[2],a[3],a[4],0,[],[])
            # R002 core, R003 editability/formula, R004 preservation, R005 structure,
            # R006 source integrity. R008 is a small anti-placeholder check.
            vals=(1,a[0],a[2],a[4],a[3],int(digest(src_path)==digest(root/"environment/input_files"/INPUT)),int(digest(_OUT_PATH)!=digest(src_path)),int(_OUT_PATH.stat().st_size>1000))
            results += [c("R002",vals[1],f"core checks={a[0]:.3f}"),c("R003",vals[2],f"formula/editable checks={a[2]:.3f}"),c("R004",vals[3],f"preservation checks={a[3]:.3f}"),c("R005",vals[4],f"special structure checks={a[4]:.3f}"),c("R006",vals[5],"input/environment hash"),c("R007",vals[6],"output differs from input"),c("R008",1-vals[7],"placeholder artifact")]
            src.close(); out.close()
        except Exception as exc:
            results += [c(k,0,f"Judge error: {exc}") for k in ("R002","R003","R004","R005","R006","R008")]; results.append(c("R007",0,"Judge error"))
    vals={x["criterion_id"]:x["score"] for x in results}; reward=max(0.0,min(1.0,sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS))); passed=vals.get("R001",0)==1 and vals.get("R002",0)>=.95 and reward>=.8
    payload={"task_id":TASK_ID,"rubric_version":TASK_ID+"-rubric-v0.1","judge_version":TASK_ID+"-judge-v0.1","criterion_results":results,"reward":round(reward,6),"passed":passed}
    args.result.parent.mkdir(parents=True,exist_ok=True); args.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); args.reward_file.parent.mkdir(parents=True,exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)
if __name__=="__main__": main()
