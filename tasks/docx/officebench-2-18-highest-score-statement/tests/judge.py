
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
    wb = load_workbook(source_dir / "score.xlsx", data_only=True, read_only=True)
    rows = list(wb.active.iter_rows(values_only=True)); wb.close()
    headers = [str(value).strip() for value in rows[0]]
    index = {name: position for position, name in enumerate(headers)}
    data = [row for row in rows[1:] if row[index["Name"]] is not None]
    best = max(data, key=lambda row: float(row[index["midterm1"]]))
    name, score = str(best[index["Name"]]), best[index["midterm1"]]
    nonempty = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    other_names = [str(row[index["Name"]]) for row in data if row != best]
    return [
        c("R002", name in text and re.search(rf"(?<!\d){re.escape(str(score))}(?!\d)", text) is not None, f"最高分学生={name}，分数={score}"),
        c("R003", any(word in text.casefold() for word in ("highest", "top score", "highest-scoring")), "明确说明该学生为最高分"),
        c("R004", len(nonempty) >= 2, "标题和奖励说明为可编辑段落"),
        c("R005", not any(other in text for other in other_names) and not re.search(r"TODO|placeholder", text, re.I), "未将其他学生误写为并列最高分"),
    ]

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--artifact-dir", type=Path, required=True); ap.add_argument("--result", type=Path, required=True); ap.add_argument("--reward-file", type=Path, required=True); a=ap.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    out = a.artifact_dir / 'highest_score_student.docx'
    results=[]
    try:
        ok = valid_doc(out); doc = Document(out) if ok else None; text = all_text(doc) if doc else ""
        results.append(c("R001", ok and (bool(text.strip()) or bool(doc and doc.tables)), "可解析的原生 OOXML DOCX"))
        results.extend(check(source_dir, out, doc, text))
        digest = hashlib.sha256((source_dir / 'score.xlsx').read_bytes()).hexdigest()
        results.append(c("R006", digest == '4a70c6310035b9b40d7a5012bdc91cc72bd539a5cd313fe7672e09e1ec65b598', f"评分所用 score.xlsx SHA256={digest}"))
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
    payload={"task_id": 'cowork-evo/officebench-2-18-highest-score-statement', "rubric_version": 'officebench-2-18-highest-score-statement'+"-v0.2", "judge_version": 'cowork-evo/officebench-2-18-highest-score-statement'+"-deterministic-v0.2", "artifact_sha256": [hashlib.sha256(out.read_bytes()).hexdigest()] if out.is_file() else [], "criterion_results": results, "reward": reward, "passed": passed, "hurdle_failures": hurdle_failures}
    a.result.parent.mkdir(parents=True, exist_ok=True); a.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    a.reward_file.parent.mkdir(parents=True, exist_ok=True); a.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)
if __name__ == "__main__": main()
