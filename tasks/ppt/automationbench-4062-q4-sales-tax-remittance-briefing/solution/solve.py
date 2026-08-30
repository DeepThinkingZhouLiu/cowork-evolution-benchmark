#!/usr/bin/env python3
from __future__ import annotations
import argparse,hashlib,shutil
from pathlib import Path
HASHES={"AutomationBench_4062_Context.json":"b597a071fc54d4432c77ca9b1b27f36163bee35402dcf6a4b2f081abd3c66717","NY_Tax_Exemption_Certificate_Notice.eml":"971213b00cc542796463ab398fd1e1890eefb44e517e7463c6bbe30932edabc8","Q4_2025_Sales_Tax_Tracker.xlsx":"d5828e994fef859802b94c709303f6d702cbd34967e958c22a39c50b826b780f","Tax_Filing_and_Payment_Authorization.eml":"0701ae39164a0c0b18a7e26cdb480a8b6dcfe6b122bb12026bb7fcca952e2f8c"}; GOLDEN_SHA="9147f69684bac9c51d73d459cabdfd4c0ea536f15b2db29441856bf4a44a2cfc"; OUTPUT_NAME="Q4_2025_Sales_Tax_Remittance_Briefing.pptx"
def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    p=argparse.ArgumentParser(); p.add_argument("--input-dir",type=Path,required=True); p.add_argument("--output-dir",type=Path,required=True); p.add_argument("--golden-dir",type=Path,required=True); a=p.parse_args()
    for n,h in HASHES.items():
        q=a.input_dir/n
        if not q.is_file() or sha(q)!=h: raise SystemExit(f"冻结输入错误：{n}")
    g=a.golden_dir/OUTPUT_NAME
    if not g.is_file() or sha(g)!=GOLDEN_SHA: raise SystemExit("Golden 错误")
    a.output_dir.mkdir(parents=True,exist_ok=True); target=a.output_dir/OUTPUT_NAME; shutil.copy2(g,target); print(target)
if __name__=="__main__": main()
