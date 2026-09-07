
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

    compact = re.sub(r"\s+", " ", text).casefold()
    exact = all(x in compact for x in ("2022-09-23 03:00:00", "glendon@gmail.com", "reminder rent due", "rent is due october 1, 2022", "rent and utilities", "rent delinquency"))
    other_subjects = any(x in compact for x in ("gradescope", "scheduled meeting", "application approved"))
    return [c("R002", exact and not other_subjects, "准确选中 Sent at 时间最早的租金提醒邮件"), c("R003", all(x in compact for x in ("date:", "subject:", "from:")), "时间、主题和发件人字段齐全"), c("R004", len([p for p in doc.paragraphs if p.text.strip()]) <= 3, "仅保留一封邮件"), c("R005", "TODO" not in text and "placeholder" not in compact, "无占位内容")]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'earliest.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'rental.eml').read_bytes()).hexdigest()
        results.append(c("R006", digest == '54f31902b3b4b18a1fefba0e4298ed8f1f2205d87595b99f22e5affa09acd807', f"输入 SHA256={digest}"))
        results.append(c("R007", bool(doc and (len(doc.paragraphs) >= 1 or doc.tables)), "输出包含可编辑段落或表格"))
    except Exception as exc:
        results = [c("R001", False, f"读取失败: {exc}")]
    vals={x["criterion_id"]: x["score"] for x in results}; reward=sum(WEIGHTS[k]*vals.get(k,0) for k in WEIGHTS)
    passed=reward >= .8 and vals.get("R001",0) == 1 and vals.get("R002",0) == 1
    payload={"task_id": 'cowork-evo/officebench-2-39-earliest-email', "rubric_version": 'officebench-2-39-earliest-email'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-39-earliest-email'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
