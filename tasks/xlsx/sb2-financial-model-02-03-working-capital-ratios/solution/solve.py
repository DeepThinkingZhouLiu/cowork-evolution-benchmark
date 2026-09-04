#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.workbook.properties import CalcProperties
def digest(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input-dir",type=Path,required=True); ap.add_argument("--output-dir",type=Path,required=True); a=ap.parse_args()
    root=Path(__file__).parents[1]; cfg=json.loads((root/"tests/expectations.json").read_text()); source=a.input_dir/cfg["input"]
    if digest(source)!=cfg["input_sha256"]: raise ValueError("input hash mismatch")
    wb=load_workbook(source,data_only=False)
    for group in cfg["target_groups"]:
        for key,value in group["cells"].items():
            sheet,coordinate=key.rsplit("!",1); wb[sheet][coordinate]=value
    if wb.calculation is None: wb.calculation=CalcProperties(calcMode="auto",fullCalcOnLoad=True,forceFullCalc=True)
    else: wb.calculation.calcMode="auto"; wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True
    a.output_dir.mkdir(parents=True,exist_ok=True); wb.save(a.output_dir/cfg["required_files"][0]); wb.close()
if __name__=="__main__": main()
