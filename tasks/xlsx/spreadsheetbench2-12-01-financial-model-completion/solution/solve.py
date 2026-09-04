#!/usr/bin/env python3
from __future__ import annotations
import argparse
from pathlib import Path
from openpyxl import load_workbook

INPUT_NAME = "financial_model_input.xlsx"
OUTPUT_NAME = "financial_model_completed.xlsx"

def solve(input_dir: Path, output_dir: Path) -> None:
    wb = load_workbook(input_dir / INPUT_NAME, data_only=False)
    ass = wb["Assumption"]
    for col, prev in zip("IJKLM", "HIJKL"):
        ass[f"{col}8"] = f"={prev}8*(1+{col}9)"
        ass[f"{col}18"] = f"={col}14*{col}19"
    wc = wb["Working Capital"]
    for col, prev in zip("IJKLM", "HIJKL"):
        wc[f"{col}6"] = f"={col}7*{col}8/365"
        wc[f"{col}8"] = 70 if col == "I" else f"={prev}8"
    for col, prev in zip("IJKLMN", "HIJKLM"):
        wc[f"{col}30"] = f"={prev}29-{col}29"
    pnl = wb["P&L"]
    for col in "CDEFGHIJKLM":
        pnl[f"{col}37"] = f"={col}36/{col}9"
    wb.calculation.fullCalcOnLoad = True; wb.calculation.forceFullCalc = True; wb.calculation.calcMode = "auto"
    output_dir.mkdir(parents=True, exist_ok=True); wb.save(output_dir / OUTPUT_NAME); wb.close()

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input-dir",type=Path,required=True); ap.add_argument("--output-dir",type=Path,required=True); a=ap.parse_args(); solve(a.input_dir,a.output_dir)
if __name__ == "__main__": main()
