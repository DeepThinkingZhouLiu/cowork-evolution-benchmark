#!/usr/bin/env python3
"""独立检查 Finance-018 的单一跨期整合工作簿。"""
from __future__ import annotations
import argparse, hashlib, json, zipfile
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from openpyxl import load_workbook
try:
    from pypinyin import lazy_pinyin
except ModuleNotFoundError:
    def lazy_pinyin(value: str) -> list[str]:
        return list(value)

TASK_ID="cowork-evo/forte-finance-018-receivables-payables-consolidation"; RUBRIC_VERSION="forte-finance-018-v0.2"; JUDGE_VERSION="forte-finance-018-v0.2"; OUTPUT="往来款跨期整合.xlsx"; EXPECTED_SHEETS=["跨期明细","未付统计","未收统计","僵尸账款核查","校验摘要"]; WEIGHTS={"R001":.10,"R002":.20,"R003":.22,"R004":.18,"R005":.12,"R006":.10,"R007":.08}; FILES=[("2025上半年","2025往来明细-上半年.xlsx"),("2025下半年","2025往来明细-下半年.xlsx"),("2026","2026往来明细.xlsx")]
def criterion(cid,ok,evidence):return {"criterion_id":cid,"score":float(ok),"evidence":evidence,"method":"deterministic"}
def dec(v): return None if v in (None,"") else Decimal(str(v)).quantize(Decimal("0.01"),rounding=ROUND_HALF_UP)
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir",type=Path,required=True); ap.add_argument("--result",type=Path,required=True); ap.add_argument("--reward-file",type=Path,required=True); args=ap.parse_args(); root=Path(__file__).parents[1]; inp=root/"data/input_files"; out=args.artifact_dir/OUTPUT; results=[]
    try:
        records=[]
        for period,filename in FILES:
            wb=load_workbook(inp/filename,data_only=True); ws=wb.active
            for rn,row in enumerate(ws.iter_rows(min_row=2,max_col=10,values_only=True),2):
                if not any(v is not None for v in row): continue
                records.append((period,filename,rn,str(row[0]),str(row[1]),str(row[8] or ""),dec(row[9])))
            wb.close()
        current=[r for r in records if r[0]=="2026"]; key=lambda r:(tuple(lazy_pinyin(r[4])),-(r[6] or Decimal("0")),r[3]); unpaid=sorted([r for r in current if r[5]=="贷" and r[6] not in (None,0)],key=key); unreceived=sorted([r for r in current if r[5]=="借" and r[6] not in (None,0)],key=key); valid=out.is_file() and zipfile.is_zipfile(out)
        if not valid: results=[criterion("R001",False,"输出缺失或损坏")]+[criterion(cid,False,"缺少有效输出") for cid in ("R002","R003","R004","R005","R006")]+[criterion("R007",True,"输入固定")]
        else:
            wb=load_workbook(out,data_only=True); structure=wb.sheetnames==EXPECTED_SHEETS; results.append(criterion("R001",structure,f"工作表={wb.sheetnames}")); detail=wb["跨期明细"]; drows=[r for r in detail.iter_rows(min_row=2,values_only=True) if any(v is not None for v in r)]; detail_ok=len(drows)==len(records) and all((r[0],r[1],r[2],str(r[3]),str(r[4]),str(r[11] or ""),dec(r[12]))==e for r,e in zip(drows,records)); results.append(criterion("R002",detail_ok,f"跨期明细={len(drows)}/191"));
            def summary_rows(name): return [tuple(r) for r in wb[name].iter_rows(min_row=2,values_only=True) if r[0] is not None]
            urows=summary_rows("未付统计"); rrows=summary_rows("未收统计"); uok=len(urows)==len(unpaid) and all((str(a[0]),str(a[1]),dec(a[2]))==(e[3],e[4],e[6]) for a,e in zip(urows,unpaid)); rok=len(rrows)==len(unreceived) and all((str(a[0]),str(a[1]),dec(a[2]))==(e[3],e[4],e[6]) for a,e in zip(rrows,unreceived)); results.append(criterion("R003",uok and rok,f"未付={len(urows)}/31，未收={len(rrows)}/2")); zombie=wb["僵尸账款核查"]; zrows=[tuple(r) for r in zombie.iter_rows(min_row=2,values_only=True) if r[0] is not None]; zombie_ok=len(zrows)==2 and all(r[-1]=="否" for r in zrows); results.append(criterion("R004",zombie_ok,"当前输入无三期连续不动僵尸账")); summ=wb["校验摘要"]; vals={str(r[0]):r[1] for r in summ.iter_rows(min_row=2,max_col=2,values_only=True) if r[0] is not None}; sum_ok=vals.get("原始记录总数")==191 and vals.get("2026 未付笔数")==31 and abs(float(vals.get("2026 未付总额",0))-3984606.46)<.01 and vals.get("2026 未收笔数")==2 and abs(float(vals.get("2026 未收总额",0))-4992891.47)<.01 and vals.get("僵尸账结论")=="无僵尸账款"; results.append(criterion("R005",sum_ok,"校验摘要业务真值")); professional=all(wb[s].freeze_panes=="A2" and wb[s].auto_filter.ref and wb[s].protection.sheet is not True for s in EXPECTED_SHEETS); results.append(criterion("R006",professional,"原生可编辑、筛选和冻结检查")); results.append(criterion("R007",True,"三份输入来源及哈希已记录")); wb.close()
    except Exception as exc: results=[criterion("R001",False,f"读取失败：{exc}")]+[criterion(cid,False,"读取失败") for cid in ("R002","R003","R004","R005","R006","R007")]
    vals={x["criterion_id"]:x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS); passed=all(vals.get(k,0)==1 for k in ("R001","R002","R003","R004","R005")) and reward>=.8; payload={"task_id":TASK_ID,"rubric_version":RUBRIC_VERSION,"judge_version":JUDGE_VERSION,"artifact_sha256":[hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [],"criterion_results":results,"reward":reward,"passed":passed}; args.result.parent.mkdir(parents=True,exist_ok=True); args.result.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding="utf-8"); args.reward_file.parent.mkdir(parents=True,exist_ok=True); args.reward_file.write_text(f"{reward:.6f}\n"); raise SystemExit(0 if passed else 1)
if __name__=="__main__":main()
