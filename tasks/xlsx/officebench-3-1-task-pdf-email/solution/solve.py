#!/usr/bin/env python3
"""Reference solution for OfficeBench 3-1/0."""
from __future__ import annotations
import argparse, hashlib, mimetypes, shutil, subprocess, tempfile
from email.message import EmailMessage
from pathlib import Path

INPUT="task_input.xlsx"; EXPECTED_SHA256="4731eca77b9807cd708f298d47d34a75b46c728929e416dc47b0ee530ee8b058"
def digest(p:Path)->str: return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input-dir",type=Path,required=True); ap.add_argument("--output-dir",type=Path,required=True); a=ap.parse_args(); src=a.input_dir/INPUT
    if digest(src)!=EXPECTED_SHA256: raise ValueError("task input hash mismatch")
    a.output_dir.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["libreoffice","--headless","--convert-to","pdf","--outdir",tmp,str(src)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
        generated=Path(tmp)/(src.stem+".pdf")
        if not generated.is_file(): raise RuntimeError("LibreOffice did not create PDF")
        pdf=a.output_dir/"task.pdf"; shutil.copyfile(generated,pdf)
    msg=EmailMessage(); msg["From"]="Alice <alice@example.com>"; msg["To"]="Alice <alice@example.com>"; msg["Subject"]="Project task sheet for review"; msg.set_content("Please review the attached project task PDF converted from the Excel task sheet.")
    msg.add_attachment((a.output_dir/"task.pdf").read_bytes(),maintype="application",subtype="pdf",filename="task.pdf")
    (a.output_dir/"Alice.eml").write_bytes(msg.as_bytes())
if __name__=="__main__": main()
