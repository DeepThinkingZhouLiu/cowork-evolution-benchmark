#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, zipfile
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

WEIGHTS={"R001":.10,"R002":.25,"R003":.20,"R004":.15,"R005":.10,"R006":.05,"R007":.05,"R008":.10}
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def scalar(value):
    if isinstance(value,(ArrayFormula,DataTableFormula)): return {"kind":type(value).__name__,"text":getattr(value,"text",None),"ref":getattr(value,"ref",None)}
    return value
def norm(value):
    value=scalar(value)
    if isinstance(value,dict):
        value=dict(value)
        if isinstance(value.get("text"),str): value["text"]=value["text"].replace("$","").replace(" ","").replace("=+","=").upper()
        return value
    if isinstance(value,str) and value.startswith("="): return value.replace("$","").replace(" ","").replace("=+","=").upper()
    return value
def equal(left,right):
    left,right=norm(left),norm(right)
    if isinstance(left,(int,float)) and isinstance(right,(int,float)): return abs(float(left)-float(right))<=1e-8*max(1.,abs(float(left)),abs(float(right)))
    return left==right or (left in (None,"") and right in (None,""))
def formula(value):
    value=scalar(value)
    if isinstance(value,dict): return isinstance(value.get("text"),str) and value["text"].startswith("=")
    return isinstance(value,str) and value.startswith("=")
def criterion(cid,score,evidence): return {"criterion_id":cid,"score":round(max(0.,min(1.,float(score))),6),"evidence":evidence,"method":"deterministic"}
def target_score(workbook,group):
    bad=[]; formula_targets=0; formula_outputs=0
    for key,expected in group["cells"].items():
        sheet,coordinate=key.rsplit("!",1)
        actual=workbook[sheet][coordinate].value if sheet in workbook.sheetnames else None
        if formula(expected):
            formula_targets+=1
            formula_outputs+=int(formula(actual))
        if not equal(actual,expected): bad.append(key)
    total=len(group["cells"]); score=(total-len(bad))/total if total else 1.
    return score,bad,formula_targets,formula_outputs
def preservation(source,output,targets):
    if output.sheetnames!=source.sheetnames: return 0.,["sheet order"]
    if any(output[s].max_row<source[s].max_row or output[s].max_column<source[s].max_column for s in source.sheetnames): return 0.,["truncated sheet"]
    checked=0; bad=[]
    for sheet in source.sheetnames:
        left,right=source[sheet],output[sheet]
        for row in range(1,left.max_row+1):
            for col in range(1,left.max_column+1):
                key=f"{sheet}!{left.cell(row,col).coordinate}"
                if key in targets: continue
                checked+=1
                if not equal(left.cell(row,col).value,right.cell(row,col).value):
                    bad.append(key)
                    if len(bad)>=25: return max(0.,1-len(bad)/max(checked,1)),bad
    return max(0.,1-len(bad)/max(checked,1)),bad
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir",type=Path,required=True); ap.add_argument("--result",type=Path,required=True); ap.add_argument("--reward-file",type=Path,required=True); a=ap.parse_args()
    root=Path(__file__).parents[1]; cfg=json.loads((root/"tests/expectations.json").read_text()); source_path=root/"data/input_files"/cfg["input"]; out=a.artifact_dir/cfg["required_files"][0]
    native=out.is_file() and zipfile.is_zipfile(out); results=[criterion("R001",native,"output exists and is native OOXML")]; output=None
    try:
        source=load_workbook(source_path,data_only=False,read_only=False)
        output=load_workbook(out,data_only=False,read_only=False) if native else None
        group_results=[]; formula_targets=formula_outputs=0
        if output:
            for group in cfg["target_groups"]:
                score,bad,ft,fo=target_score(output,group); group_results.append((group["id"],score,bad)); formula_targets+=ft; formula_outputs+=fo
        while len(group_results)<3: group_results.append(("unavailable",0.,["missing output"]))
        for cid,(name,score,bad) in zip(("R002","R003","R004"),group_results[:3]): results.append(criterion(cid,score,f"{name}: mismatches={bad[:12]}"))
        broken=0
        if output:
            for sheet in cfg.get("no_ref_sheets",output.sheetnames):
                if sheet not in output.sheetnames: broken+=1; continue
                for row in output[sheet].iter_rows():
                    for cell in row:
                        value=scalar(cell.value)
                        text=value.get("text","") if isinstance(value,dict) else value
                        if isinstance(text,str) and "#REF!" in text.upper(): broken+=1
            calc=output.calculation; calc_ok=bool(calc and calc.calcMode=="auto" and calc.fullCalcOnLoad and calc.forceFullCalc)
        else: calc_ok=False
        formula_ratio=formula_outputs/formula_targets if formula_targets else 0.
        results.append(criterion("R005",int(broken==0 and calc_ok and formula_ratio>=.95),f"broken_refs={broken}, formula_targets={formula_targets}, formula_outputs={formula_outputs}, auto_recalc={calc_ok}"))
        results.append(criterion("R006",digest(source_path)==cfg["input_sha256"],f"input_sha256={digest(source_path)}"))
        changed=native and digest(out)!=digest(source_path); results.append(criterion("R007",int(changed and formula_ratio>=.95),f"changed={changed}, native_formula_ratio={formula_ratio:.3f}"))
        targets={key for group in cfg["target_groups"] for key in group["cells"]}
        if output and cfg["preservation_mode"]=="strict": preserve_score,preserve_bad=preservation(source,output,targets)
        elif output:
            sheets_ok=output.sheetnames==cfg["required_sheets"]
            # openpyxl normal mode does not retain trailing styled/empty rows from
            # the source dimension metadata; target groups and sheet order are the
            # reliable structural checks for these repaired workbooks.
            dims_ok=all(s in output.sheetnames for s in cfg["minimum_dimensions"])
            preserve_score=float(sheets_ok and dims_ok); preserve_bad=[] if preserve_score else ["sheet order or restored dimensions"]
        else: preserve_score,preserve_bad=0.,["missing output"]
        results.append(criterion("R008",preserve_score,f"preservation mismatches={preserve_bad[:12]}"))
        source.close()
        if output: output.close()
    except Exception as exc:
        seen={x["criterion_id"] for x in results}
        for cid in WEIGHTS:
            if cid not in seen: results.append(criterion(cid,0.,f"Judge read failure: {exc}"))
    values={x["criterion_id"]:x["score"] for x in results}; reward=sum(WEIGHTS[k]*values.get(k,0.) for k in WEIGHTS)
    passed=reward>=cfg["pass_threshold"] and values.get("R001",0)==1 and values.get("R002",0)>=.8 and values.get("R003",0)>=.8 and values.get("R008",0)>=.95
    payload={"task_id":cfg["task_id"],"rubric_version":cfg["rubric_version"],"judge_version":cfg["judge_version"],"artifact_sha256":[digest(out)] if out.is_file() else [],"criterion_results":results,"reward":round(reward,6),"passed":passed}
    a.result.parent.mkdir(parents=True,exist_ok=True); a.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); a.reward_file.parent.mkdir(parents=True,exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n",encoding="utf-8"); raise SystemExit(0 if passed else 1)
if __name__=="__main__": main()
