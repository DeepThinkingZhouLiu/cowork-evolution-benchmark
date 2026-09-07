
from __future__ import annotations
import argparse, hashlib, json, re, zipfile
from collections import Counter
from pathlib import Path
from docx import Document
from openpyxl import load_workbook

WEIGHTS = {"R001": .20, "R002": .20, "R003": .16, "R004": .14, "R005": .15, "R006": .10, "R007": .05}
def c(cid, ok, evidence): return {"criterion_id": cid, "score": 1.0 if ok else 0.0, "evidence": evidence, "method": "deterministic"}
def valid_doc(path): return path.is_file() and zipfile.is_zipfile(path)
def all_text(doc): return "\n".join(p.text for p in doc.paragraphs)
def check(source_dir, out, doc, text):

    wb=load_workbook(source_dir / "CS161_Class_Roster.xlsx", data_only=True, read_only=True); ws=wb.active; headers=[str(x).strip() for x in next(ws.iter_rows(values_only=True))]; ix={h:i for i,h in enumerate(headers)}; expected=[f"{r[ix['Student ID']]}, {r[ix['First Name']]} {r[ix['Last Name']]}" for r in ws.iter_rows(min_row=2,values_only=True) if str(r[ix['Major']]).strip()=="CS"]; wb.close(); actual=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return [c("R002", actual == expected, "CS 专业学生和顺序准确"), c("R003", len(actual) == len(expected), "匹配行数准确"), c("R004", all(", " in x for x in actual), "格式为 ID, name"), c("R005", "Math" not in text and "ECE" not in text, "未混入其他专业")]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'cs_class_2024.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'CS161_Class_Roster.xlsx').read_bytes()).hexdigest()
        results.append(c("R006", digest == 'c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7', f"输入 SHA256={digest}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    passed=reward >= .8 and vals.get("R001",0) == 1 and vals.get("R002",0) == 1
    payload={"task_id": 'cowork-evo/officebench-2-20-cs-major-list', "rubric_version": 'officebench-2-20-cs-major-list'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-20-cs-major-list'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
