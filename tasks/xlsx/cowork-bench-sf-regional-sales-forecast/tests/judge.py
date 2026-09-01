#!/usr/bin/env python3
"""独立检查单一区域预测工作簿及其俄文说明表。"""
from __future__ import annotations
import argparse, hashlib, json, re, zipfile
from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook

TASK_ID="cowork-evo/cowork-bench-sf-regional-sales-forecast"; RUBRIC_VERSION="cowork-bench-sf-regional-sales-static-v0.2"; JUDGE_VERSION="cowork-bench-sf-regional-sales-static-v0.2"; OUTPUT="Sales_Forecast.xlsx"
REGIONS=["Азиатско-Тихоокеанский регион","Европа","Латинская Америка","Ближний Восток","Северная Америка"]
HIST_HEADERS=["Region","Month","Orders","Revenue"]; FORECAST_HEADERS=["Region","Trailing_3M_Avg","GDP_Growth_Pct","Economic_Adjustment","Adjusted_Monthly_Forecast","Next_Quarter_Forecast"]
SUMMARY_HEADERS=["Metric","Value"]; WEIGHTS={"R001":.10,"R002":.18,"R003":.16,"R004":.18,"R005":.16,"R006":.10,"R007":.07,"R008":.05}
def criterion(cid,ok,evidence): return {"criterion_id":cid,"score":float(ok),"evidence":evidence,"method":"deterministic"}
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def close(a,b,tol=.011):
    try:return abs(float(a)-float(b))<=tol
    except (TypeError,ValueError):return False
def norm(v):return re.sub(r"\s+"," ",str(v or "")).strip().lower()
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir",type=Path,required=True); ap.add_argument("--result",type=Path,required=True); ap.add_argument("--reward-file",type=Path,required=True); args=ap.parse_args(); root=Path(__file__).parents[1]; inp=root/"data/input_files"; out=args.artifact_dir/OUTPUT; results=[]
    try:
        hist=json.loads((inp/"historical_sales.json").read_text(encoding="utf-8")); econ=json.loads((inp/"economic_indicators.json").read_text(encoding="utf-8"))["regions"]; groups=defaultdict(list); lookup={}
        for row in hist: groups[row["region"]].append(float(row["revenue"])); lookup[(row["region"],row["month"])]=row
        expected={};
        for r in REGIONS:
            raw=sum(groups[r])/3; trail=round(raw,2); gdp=float(econ[r]["gdp_growth"]); adj=round(1+gdp/100,4); monthly=round(trail*adj,2); quarter=round(monthly*3,2); expected[r]=(trail,gdp,adj,monthly,quarter,raw*adj*3)
        total_q4=round(sum(float(r["revenue"]) for r in hist),2); total_q1=round(sum(round(x[-1],2) for x in expected.values()),2); high=max(expected,key=lambda r:expected[r][-1]); low=min(expected,key=lambda r:expected[r][-1]); high_gdp=max(expected,key=lambda r:expected[r][1])
        summary={"Total_Q4_2025_Revenue":total_q4,"Total_Monthly_Average":round(total_q4/3,2),"Total_Q1_2026_Forecast":total_q1,"QoQ_Growth_Pct":round((total_q1/total_q4-1)*100,1),"Highest_Forecast_Region":high,"Highest_Forecast_Amount":round(expected[high][-1],2),"Lowest_Forecast_Region":low,"Lowest_Forecast_Amount":round(expected[low][-1],2),"Highest_GDP_Growth_Region":high_gdp,"Regions_Count":5}
        valid=out.is_file() and zipfile.is_zipfile(out)
        if not valid:
            results=[criterion("R001",False,"输出缺失或损坏")]+[criterion(cid,False,"缺少有效输出") for cid in ("R002","R003","R004","R005","R006")]+[criterion("R007",False,"缺少有效输出"),criterion("R008",True,"输入哈希由作者侧固定")]
        else:
            wb=load_workbook(out,data_only=True); structure=wb.sheetnames==["Historical","Forecast","Summary","Forecast Notes"]; results.append(criterion("R001",structure,f"工作表={wb.sheetnames}"))
            hs=wb["Historical"]; hrows=list(hs.iter_rows(min_row=2,values_only=True)); h_ok=tuple(x.value for x in hs[1][:4])==tuple(HIST_HEADERS) and len(hrows)==15 and all((str(row[0]),str(row[1])) in lookup and close(row[2],lookup[(str(row[0]),str(row[1]))]["orders"],.001) and close(row[3],lookup[(str(row[0]),str(row[1]))]["revenue"],.011) for row in hrows); results.append(criterion("R002",h_ok,f"Historical 行数={len(hrows)}"))
            fs=wb["Forecast"]; frows=list(fs.iter_rows(min_row=2,values_only=True)); factual={str(row[0]):row for row in frows}; t_ok=a_ok=m_ok=len(frows)==5 and set(factual)==set(REGIONS)
            for r in REGIONS:
                row=factual.get(r); exp=expected[r]
                if row is None: t_ok=a_ok=m_ok=False; continue
                t_ok &= close(row[1],exp[0]); a_ok &= close(row[2],exp[1],.0001) and close(row[3],exp[2],.00011); m_ok &= close(row[4],exp[3]) and close(row[5],exp[4])
            results.extend([criterion("R003",t_ok,"移动平均检查"),criterion("R004",a_ok,"GDP 和调整系数检查"),criterion("R005",m_ok,"月度/季度预测检查")])
            ss=wb["Summary"]; actual={str(row[0]):row[1] for row in ss.iter_rows(min_row=2,max_col=2,values_only=True) if row[0] is not None}; s_ok=len(actual)==10 and all((norm(actual.get(k))==norm(v) if isinstance(v,str) else close(actual.get(k),v,.11 if k=="QoQ_Growth_Pct" else .011)) for k,v in summary.items()); results.append(criterion("R006",s_ok,f"Summary 指标={len(actual)}/10"))
            ns=wb["Forecast Notes"]; note_text=" ".join(str(x or "") for row in ns.iter_rows(values_only=True) for x in row); sections=["краткое резюме","методолог","региональн","допущен","рекомендац"]; notes_ok=all(x in norm(note_text) for x in sections) and all(norm(r) in norm(note_text) for r in REGIONS); results.append(criterion("R007",notes_ok,"俄文说明章节和地区检查"))
            professional=all(wb[s].freeze_panes=="A2" and wb[s].auto_filter.ref and wb[s].protection.sheet is not True for s in wb.sheetnames); results.append(criterion("R008",professional,"各表筛选、冻结和可编辑检查")); wb.close()
    except Exception as exc:
        results=[criterion("R001",False,f"读取失败：{exc}")]+[criterion(cid,False,"读取失败") for cid in ("R002","R003","R004","R005","R006","R007","R008")]
    vals={x["criterion_id"]:x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS); passed=all(vals.get(k,0)==1 for k in ("R001","R002","R003","R004","R005","R006","R007")) and reward>=.75; payload={"task_id":TASK_ID,"rubric_version":RUBRIC_VERSION,"judge_version":JUDGE_VERSION,"artifact_sha256":[digest(out)] if out.is_file() else [],"criterion_results":results,"reward":reward,"passed":passed}; args.result.parent.mkdir(parents=True,exist_ok=True); args.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); args.reward_file.parent.mkdir(parents=True,exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)
if __name__=="__main__": main()
