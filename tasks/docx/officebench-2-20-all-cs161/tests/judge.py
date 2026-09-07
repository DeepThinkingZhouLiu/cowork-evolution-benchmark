
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
    wb=load_workbook(source_dir / "CS161_Class_Roster.xlsx", data_only=True, read_only=True)
    rows=list(wb.active.iter_rows(values_only=True)); wb.close()
    headers=[str(x).strip() for x in rows[0]]; ix={h:i for i,h in enumerate(headers)}
    expected=[f"{r[ix['Student ID']]}, {r[ix['First Name']]} {r[ix['Last Name']]}" for r in rows[1:] if r[ix['Student ID']] is not None]
    actual=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return [
        c("R002", actual == expected, f"学生名单及源顺序完全匹配 {len(expected)} 行"),
        c("R003", len(actual) == len(expected), f"学生数量={len(actual)}/{len(expected)}"),
        c("R004", all(re.fullmatch(r"S\d{4}, .+ .+", item) for item in actual), "每人独立段落且格式为 Student ID, First Name Last Name"),
        c("R005", len(set(actual)) == len(actual) and not re.search(r"TODO|placeholder", text, re.I), "无重复行或占位内容"),
    ]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'students.docx'
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
    hurdle_failures = [key for key in ("R001", "R002") if vals.get(key, 0) != 1]
    if "R001" in hurdle_failures:
        reward = 0.0
    elif hurdle_failures:
        reward = min(reward, 0.65)
    passed=reward >= .8 and not hurdle_failures
    payload={"task_id": 'cowork-evo/officebench-2-20-all-cs161', "rubric_version": 'officebench-2-20-all-cs161'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-20-all-cs161'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed, "hurdle_failures": hurdle_failures}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
