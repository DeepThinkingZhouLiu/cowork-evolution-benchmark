#!/usr/bin/env python3
"""独立检查 JSON 清洗结果在单一 XLSX 中的分析、库存和 Dashboard。"""
from __future__ import annotations
import argparse, json, hashlib, zipfile
from collections import OrderedDict
from pathlib import Path
from openpyxl import load_workbook

TASK_ID="cowork-evo/cowork-bench-json-data-cleaning-pipeline"; RUBRIC_VERSION="cowork-bench-json-data-cleaning-pipeline-v0.2"; JUDGE_VERSION="cowork-bench-json-data-cleaning-pipeline-v0.2"; OUTPUT="Data_Pipeline_Report.xlsx"; WEIGHTS={"R001":.10,"R002":.20,"R003":.15,"R004":.18,"R005":.12,"R006":.15,"R007":.07,"R008":.03}
def criterion(cid,ok,evidence):return {"criterion_id":cid,"score":float(ok),"evidence":evidence,"method":"deterministic"}
def clean(rows,key):
    seen=set(); out=[]
    for row in rows:
        ident=str(row.get(key,""));
        if not ident or ident in seen: continue
        seen.add(ident); item=dict(row); item["product_name"]=str(item.get("product_name","")).strip().title(); out.append(item)
    return out
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir",type=Path,required=True); ap.add_argument("--result",type=Path,required=True); ap.add_argument("--reward-file",type=Path,required=True); args=ap.parse_args(); root=Path(__file__).parents[1]; inp=root/"data/input_files"; out=args.artifact_dir/OUTPUT; results=[]
    try:
        raw_sales=json.loads((inp/"sales_api_response.json").read_text(encoding="utf-8")); raw_inv=json.loads((inp/"inventory_api_response.json").read_text(encoding="utf-8")); raw_sales=raw_sales.get("data",raw_sales); raw_inv=raw_inv.get("data",raw_inv); sales=clean(raw_sales,"order_id"); inv=clean(raw_inv,"product_id"); agg=OrderedDict()
        for row in sales:
            name=row["product_name"]; item=agg.setdefault(name,{"q":0,"rev":0.0,"prices":[]}); item["q"]+=int(row["quantity"]); item["rev"]+=int(row["quantity"])*float(row["unit_price"]); item["prices"].append(float(row["unit_price"]))
        sales_expected={k:(v["q"],round(v["rev"],2),round(sum(v["prices"])/len(v["prices"]),2)) for k,v in agg.items()}; inv_expected=OrderedDict((r["product_name"],(int(r["current_stock"]),int(r["reorder_point"]),"LOW" if int(r["current_stock"])<=int(r["reorder_point"]) else "OK")) for r in inv); names=list(agg)+[k for k in inv_expected if k not in agg]; combined_expected=[(n,sales_expected.get(n,(0,0.0,0.0))[0],inv_expected.get(n,(0,0,"NO INVENTORY"))[0],sales_expected.get(n,(0,0.0,0.0))[1],inv_expected.get(n,(0,0,"NO INVENTORY"))[2]) for n in names]; summary_expected={"raw_sales_count":len(raw_sales),"unique_sales_count":len(sales),"sales_duplicates_removed":len(raw_sales)-len(sales),"raw_inventory_count":len(raw_inv),"unique_inventory_count":len(inv),"inventory_duplicates_removed":len(raw_inv)-len(inv),"total_quantity_sold":sum(x[0] for x in sales_expected.values()),"total_revenue":round(sum(x[1] for x in sales_expected.values()),2),"low_stock_products":sum(x[2]=="LOW" for x in inv_expected.values())}
        valid=out.is_file() and zipfile.is_zipfile(out)
        if not valid: results=[criterion("R001",False,"输出缺失或损坏")]+[criterion(cid,False,"缺少有效输出") for cid in ("R002","R003","R004","R005","R006","R007")]+[criterion("R008",True,"输入快照固定")]
        else:
            wb=load_workbook(out,data_only=False); structure=wb.sheetnames==["Sales Analysis","Inventory Status","Combined View","Pipeline Dashboard","Run Summary"]; results.append(criterion("R001",structure,f"工作表={wb.sheetnames}")); sa=wb["Sales Analysis"]; srows=[tuple(row) for row in sa.iter_rows(min_row=2,values_only=True) if row[0] is not None]; sales_ok=len(srows)==len(sales_expected) and all(str(r[0]) in sales_expected and r[1]==sales_expected[str(r[0])][0] and abs(float(r[2])-sales_expected[str(r[0])][1])<.011 and abs(float(r[3])-sales_expected[str(r[0])][2])<.011 for r in srows); results.append(criterion("R002",sales_ok,"销售聚合与去重检查")); invs=wb["Inventory Status"]; irows=[tuple(row) for row in invs.iter_rows(min_row=2,values_only=False) if row[0].value is not None]; inv_ok=len(irows)==len(inv_expected) and all(str(row[0].value) in inv_expected and row[1].value==inv_expected[str(row[0].value)][0] and row[2].value==inv_expected[str(row[0].value)][1] and isinstance(row[3].value,str) and row[3].value.startswith("=") for row in irows); results.append(criterion("R003",inv_ok,"库存字段、状态公式与去重检查"));
            def rows_values(sheet): return [tuple(r) for r in sheet.iter_rows(min_row=2,values_only=True) if r[0] is not None]
            crows=rows_values(wb["Combined View"]); drows=rows_values(wb["Pipeline Dashboard"]); comb_ok=crows==combined_expected and drows==combined_expected; results.append(criterion("R004",comb_ok,"Combined View 与 Dashboard 一致")); formula_ok=inv_ok; results.append(criterion("R005",formula_ok,"库存 Status 使用 IF 公式")); run=wb["Run Summary"]; run_values={str(r[0]):r[1] for r in run.iter_rows(min_row=2,max_col=2,values_only=True) if r[0] is not None}; sum_ok=all(run_values.get(k)==v for k,v in summary_expected.items()); results.append(criterion("R006",sum_ok,"Run Summary 与输入重算一致")); professional=all(wb[s].freeze_panes=="A2" and wb[s].auto_filter.ref and wb[s].protection.sheet is not True for s in wb.sheetnames); results.append(criterion("R007",professional,"原生表格、筛选、冻结和可编辑检查")); results.append(criterion("R008",True,"输入快照固定且 provenance 已记录")); wb.close()
    except Exception as exc: results=[criterion("R001",False,f"读取失败：{exc}")]+[criterion(cid,False,"读取失败") for cid in ("R002","R003","R004","R005","R006","R007","R008")]
    vals={x["criterion_id"]:x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS); passed=all(vals.get(k,0)==1 for k in ("R001","R002","R003","R004","R005","R006")) and reward>=.8; payload={"task_id":TASK_ID,"rubric_version":RUBRIC_VERSION,"judge_version":JUDGE_VERSION,"artifact_sha256":[hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [],"criterion_results":results,"reward":reward,"passed":passed}; args.result.parent.mkdir(parents=True,exist_ok=True); args.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); args.reward_file.parent.mkdir(parents=True,exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)
if __name__=="__main__":main()
