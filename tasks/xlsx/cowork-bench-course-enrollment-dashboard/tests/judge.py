#!/usr/bin/env python3
"""独立确定性 Judge：从 Canvas 投影输入重算课程指标并检查 XLSX。"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import subprocess
import tempfile
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook


OUTPUT = "Course_Enrollment_Analytics.xlsx"
SHEETS = ["Course_Metrics", "Risk_Distribution", "Submission_Quality", "Risk_Action_Plan", "Source_Notes"]
COURSE_IDS = [16, 17]
RISK_ORDER = ["Critical", "At Risk", "On Track"]
INPUT_HASHES = {
    "canvas_courses.csv": "72f17c5402df529b0d8c822fe87050c4c3c75c5f7c220d963fb592d957557162",
    "canvas_enrollments.csv": "afb3b26cb6a101d741c3cf3560b03ccc2f17047b79fc70b01b7864f9957cad3e",
    "canvas_assignments.csv": "98573aae81792ae965503241dff2a162a9170d20d86daed8a337c72e765d6d16",
    "canvas_submissions.csv": "11f34a9c4fe74934b1f8c9576358f94d45a7ae63f60efa66f02365ec0b6c2125",
    "Retention_Policy.pdf": "1e851c80117c065d3535716b98b4486e41e3b6a7a7372662947a77f501baa4d8",
    "scoring_model.json": "ef019f895d416c3495c59d43e867125aea63b7f4fda929e2959771a980483c9e",
}
SPECS = {
    "R001": ("工作簿存在、可解析且只包含指定输出", 10, "hurdle"),
    "R002": ("Course_Metrics 的课程和选课事实准确", 16, "hurdle"),
    "R003": ("Course_Metrics 的提交、完成率和迟交率准确", 15, "hurdle"),
    "R004": ("平均分、排序和风险级别准确", 15, "hurdle"),
    "R005": ("Risk_Distribution 原生公式重算正确", 12, "hurdle"),
    "R006": ("Submission_Quality 与课程指标一致且可编辑", 10, "hurdle"),
    "R007": ("Risk_Action_Plan 逐项保留规则", 8, "hurdle"),
    "R008": ("Source_Notes 含来源、口径和输入哈希", 5, "positive"),
    "R009": ("冻结首行、原生表格、加粗表头和合理列宽", 5, "positive"),
    "R010": ("工作簿由 LibreOffice 重算并可读", 5, "positive"),
    "R011": ("输出不包含截图、额外文件或空表", 4, "hurdle"),
    "R012": ("风险阈值边界和三类覆盖完整", 5, "positive"),
}
CAPS = {"R001": 0.0, "R002": 0.35, "R003": 0.45, "R004": 0.5, "R005": 0.55, "R006": 0.6, "R007": 0.65, "R011": 0.0}


def result(cid: str, passed: bool, evidence: str = "") -> dict:
    description, weight, kind = SPECS[cid]
    return {"criterion_id": cid, "description": description, "weight": weight, "type": kind, "passed": bool(passed), "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic"}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def rows(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as stream:
        return list(csv.DictReader(stream))


def val_number(value) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(str(value).replace(",", ".").replace("%", "").strip())
    except ValueError:
        return None


def truthy(value) -> bool:
    return str(value or "").strip().casefold() in {"t", "true", "1", "yes", "y"}


def derive(source: Path) -> dict:
    courses = {int(r["course_id"]): r for r in rows(source / "canvas_courses.csv")}
    enrollments = rows(source / "canvas_enrollments.csv")
    assignments = rows(source / "canvas_assignments.csv")
    assignment_course = {int(r["assignment_id"]): int(r["course_id"]) for r in assignments}
    model = json.loads((source / "scoring_model.json").read_text(encoding="utf-8"))
    high_cutoff = float(next(x for x in model["risk_classification"]["high_risk"]["threshold"].split() if x.replace(".", "", 1).isdigit()))
    low_cutoff = 75.0
    by_course: dict[int, list[dict[str, str]]] = defaultdict(list)
    score_by_course: dict[int, list[float]] = defaultdict(list)
    scored_users: dict[int, set[str]] = defaultdict(set)
    for r in enrollments:
        cid = int(r["course_id"])
        if cid not in COURSE_IDS:
            continue
        by_course[cid].append(r)
        score = val_number(r.get("current_score"))
        if "student" in r.get("type", "").casefold() and score is not None:
            score_by_course[cid].append(score)
            scored_users[cid].add(str(r["user_id"]))
    submissions = rows(source / "canvas_submissions.csv")
    sub_count: Counter[int] = Counter()
    graded_count: Counter[int] = Counter()
    late_count: Counter[int] = Counter()
    for r in submissions:
        cid = assignment_course.get(int(r["assignment_id"]))
        if cid not in COURSE_IDS:
            continue
        sub_count[cid] += 1
        if r.get("workflow_state", "").casefold() == "graded":
            graded_count[cid] += 1
        if truthy(r.get("late")):
            late_count[cid] += 1
    metrics = []
    for cid in COURSE_IDS:
        scores = score_by_course[cid]
        total = sub_count[cid]
        avg = round(sum(scores) / len(scores), 2)
        risk = "Critical" if avg < high_cutoff else "At Risk" if avg < low_cutoff else "On Track"
        metrics.append({
            "Course_ID": cid,
            "Course_Name": courses[cid]["course_name"],
            "Course_Code": courses[cid]["course_code"],
            "Total_Enrollment": len(by_course[cid]),
            "Student_Enrollment": sum("student" in r.get("type", "").casefold() for r in by_course[cid]),
            "Graded_Students": len(scored_users[cid]),
            "Submission_Count": total,
            "Graded_Submissions": graded_count[cid],
            "Avg_Score": avg,
            "Completion_Rate": round(100 * graded_count[cid] / total, 1) if total else 0.0,
            "Late_Rate": round(100 * late_count[cid] / total, 1) if total else 0.0,
            "Late_Submissions": late_count[cid],
            "Risk_Level": risk,
        })
    metrics.sort(key=lambda r: (r["Avg_Score"], r["Course_ID"]))
    distribution = [{
        "Risk_Level": risk,
        "Course_Count": sum(row["Risk_Level"] == risk for row in metrics),
        "Student_Enrollment": sum(row["Student_Enrollment"] for row in metrics if row["Risk_Level"] == risk),
    } for risk in RISK_ORDER]
    actions = model["intervention_actions"]
    action_rows = [
        {"Risk_Level": "Critical", "Action": actions["high"]["action"], "Timeline": actions["high"]["timeline"], "Responsible": actions["high"]["responsible"]},
        {"Risk_Level": "At Risk", "Action": actions["medium"]["action"], "Timeline": actions["medium"]["timeline"], "Responsible": actions["medium"]["responsible"]},
        {"Risk_Level": "On Track", "Action": actions["low"]["action"], "Timeline": actions["low"]["timeline"], "Responsible": actions["low"]["responsible"]},
    ]
    return {"metrics": metrics, "distribution": distribution, "actions": action_rows}


def nonempty_rows(ws) -> list[tuple]:
    return [tuple(row) for row in ws.iter_rows(min_row=2, values_only=True) if any(v is not None and str(v).strip() for v in row)]


def recalc(path: Path) -> Path | None:
    binary = shutil.which("libreoffice") or shutil.which("soffice")
    if not binary:
        return None
    root = Path(tempfile.mkdtemp(prefix="cowork-course-recalc-"))
    source = root / "source"; output = root / "output"
    source.mkdir(); output.mkdir()
    copy = source / OUTPUT
    shutil.copy2(path, copy)
    process = subprocess.run([binary, "--headless", f"-env:UserInstallation={(root / 'profile').resolve().as_uri()}", "--convert-to", "xlsx", "--outdir", str(output), str(copy)], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=300, check=False)
    result_path = output / OUTPUT
    return result_path if process.returncode == 0 and result_path.is_file() else None


def evaluate(artifact_dir: Path) -> dict:
    source = Path(__file__).resolve().parents[1] / "data" / "input_files"
    artifact_dir = artifact_dir.resolve()
    output = artifact_dir / OUTPUT
    checks: list[dict] = []
    names = sorted(p.name for p in artifact_dir.iterdir()) if artifact_dir.is_dir() else []
    valid = output.is_file() and zipfile.is_zipfile(output)
    formula_wb = None
    try:
        if valid:
            formula_wb = load_workbook(output, data_only=False)
            valid = formula_wb.sheetnames == SHEETS
    except Exception as error:
        valid = False
        formula_wb = None
        parse_error = str(error)
    else:
        parse_error = ""
    checks.append(result("R001", valid and names == [OUTPUT], f"valid={valid}; files={names}; sheets={formula_wb.sheetnames if formula_wb else None}; error={parse_error}"))
    expected = derive(source)
    if not valid:
        checks.extend([result(cid, False, "工作簿无效") for cid in ("R002", "R003", "R004", "R005", "R006", "R007")])
        checks.extend([result("R008", False, "无法读取来源说明"), result("R009", False, "无法读取格式"), result("R010", False, "无法重算"), result("R011", False, "无有效输出"), result("R012", False, "无有效输出")])
        return finalize(checks)
    value_path = recalc(output)
    value_wb = load_workbook(value_path, data_only=True) if value_path else load_workbook(output, data_only=True)
    metric_headers = [c.value for c in formula_wb["Course_Metrics"][1]]
    metric_rows = nonempty_rows(value_wb["Course_Metrics"])
    exp_headers = ["Course_ID", "Course_Name", "Course_Code", "Total_Enrollment", "Student_Enrollment", "Graded_Students", "Submission_Count", "Graded_Submissions", "Avg_Score", "Completion_Rate", "Late_Rate", "Late_Submissions", "Risk_Level"]
    metric_ok = metric_headers == exp_headers and len(metric_rows) == len(expected["metrics"])
    if metric_ok:
        for actual, exp in zip(metric_rows, expected["metrics"]):
            for key, idx in [("Course_ID", 0), ("Course_Name", 1), ("Course_Code", 2), ("Total_Enrollment", 3), ("Student_Enrollment", 4), ("Graded_Students", 5), ("Submission_Count", 6), ("Graded_Submissions", 7), ("Avg_Score", 8), ("Completion_Rate", 9), ("Late_Rate", 10), ("Late_Submissions", 11), ("Risk_Level", 12)]:
                if isinstance(exp[key], (int, float)):
                    metric_ok &= actual[idx] is not None and abs(float(actual[idx]) - float(exp[key])) <= (0.011 if key in {"Avg_Score", "Completion_Rate", "Late_Rate"} else 0.001)
                else:
                    metric_ok &= str(actual[idx]) == str(exp[key])
    checks.append(result("R002", metric_ok, f"rows={len(metric_rows)}; headers={metric_headers}"))
    submission_ok = metric_ok
    if submission_ok:
        submission_ok = all(row[6] >= row[7] and 0 <= row[9] <= 100 and 0 <= row[10] <= 100 for row in metric_rows)
    checks.append(result("R003", submission_ok, "提交、完成率、迟交率与独立重算一致"))
    avg_ok = metric_ok and [float(row[8]) for row in metric_rows] == sorted(float(row[8]) for row in metric_rows)
    checks.append(result("R004", avg_ok, f"avgs={[row[8] for row in metric_rows]}"))

    dist_ws = formula_wb["Risk_Distribution"]
    dist_values = nonempty_rows(value_wb["Risk_Distribution"])
    dist_headers = [c.value for c in dist_ws[1]]
    dist_expected = expected["distribution"]
    dist_ok = dist_headers == ["Risk_Level", "Course_Count", "Student_Enrollment"] and len(dist_values) == 3
    if dist_ok:
        for actual, exp in zip(dist_values, dist_expected):
            dist_ok &= actual[0] == exp["Risk_Level"] and float(actual[1]) == exp["Course_Count"] and float(actual[2]) == exp["Student_Enrollment"]
    formulas = [dist_ws[f"B{i}"].value for i in range(2, 5)] + [dist_ws[f"C{i}"].value for i in range(2, 5)]
    formula_shape = all(isinstance(f, str) and f.startswith("=") and ("COUNTIF" in f.upper() or "SUMIF" in f.upper()) for f in formulas)
    checks.append(result("R005", dist_ok and formula_shape, f"values={dist_values}; formulas={formulas}"))

    q_ws = formula_wb["Submission_Quality"]
    q_values = nonempty_rows(value_wb["Submission_Quality"])
    q_headers = [c.value for c in q_ws[1]]
    q_ok = q_headers == ["Course_Name", "Submission_Count", "Graded_Submissions", "Completion_Rate", "Late_Submissions", "Late_Rate"] and len(q_values) == 2
    if q_ok:
        for actual, exp in zip(q_values, expected["metrics"]):
            q_ok &= actual[0] == exp["Course_Name"] and float(actual[1]) == exp["Submission_Count"] and float(actual[2]) == exp["Graded_Submissions"] and abs(float(actual[3]) - exp["Completion_Rate"]) <= 0.011 and float(actual[4]) == exp["Late_Submissions"] and abs(float(actual[5]) - exp["Late_Rate"]) <= 0.011
    quality_formulas = [q_ws.cell(row=r, column=c).value for r in range(2, 4) for c in range(1, 7)]
    checks.append(result("R006", q_ok and all(isinstance(f, str) and f.startswith("=") for f in quality_formulas), f"values={q_values}"))

    action_values = nonempty_rows(value_wb["Risk_Action_Plan"])
    action_ok = [c.value for c in formula_wb["Risk_Action_Plan"][1]] == ["Risk_Level", "Action", "Timeline", "Responsible"] and action_values == [(x["Risk_Level"], x["Action"], x["Timeline"], x["Responsible"]) for x in expected["actions"]]
    checks.append(result("R007", action_ok, f"rows={action_values}"))

    notes = " ".join(str(v or "") for row in nonempty_rows(value_wb["Source_Notes"]) for v in row)
    notes_ok = all(token in notes for token in ("Cowork-Bench", "d943e75bc0fc8e3b27141979300cd8cbcd1e890d", "Critical", "current_score", "canvas_courses.csv")) and all(digest(source / name) in notes for name in INPUT_HASHES)
    checks.append(result("R008", notes_ok, f"notes_len={len(notes)}"))

    formatting = {
        "tables": sum(len(ws.tables) for ws in formula_wb.worksheets),
        "freezes": {ws.title: ws.freeze_panes for ws in formula_wb.worksheets},
        "bold_headers": all(bool(ws["A1"].font.bold) for ws in formula_wb.worksheets),
        "widths": all(any((dim.width or 0) >= 12 for dim in ws.column_dimensions.values()) for ws in formula_wb.worksheets),
    }
    checks.append(result("R009", formatting["tables"] >= 5 and all(v == "A2" for v in formatting["freezes"].values()) and formatting["bold_headers"] and formatting["widths"], json.dumps(formatting)))
    checks.append(result("R010", value_path is not None and value_wb.sheetnames == SHEETS, f"recalculated={value_path is not None}"))
    checks.append(result("R011", names == [OUTPUT] and all(ws.max_row >= 2 for ws in formula_wb.worksheets), f"files={names}"))
    checks.append(result("R012", [x["Risk_Level"] for x in expected["distribution"]] == RISK_ORDER and {x["Risk_Level"] for x in expected["metrics"]}.issubset(set(RISK_ORDER)), "三类风险顺序完整"))
    return finalize(checks)


def finalize(checks: list[dict]) -> dict:
    for cid in SPECS:
        if not any(c["criterion_id"] == cid for c in checks):
            checks.append(result(cid, False, "未执行"))
    checks.sort(key=lambda c: c["criterion_id"])
    positive_total = sum(spec[1] for spec in SPECS.values() if spec[1] > 0)
    earned = sum(c["weight"] for c in checks if c["weight"] > 0 and c["passed"])
    reward = max(0.0, min(1.0, earned / positive_total))
    failures = [c["criterion_id"] for c in checks if c["type"] == "hurdle" and not c["passed"]]
    if failures:
        reward = min(reward, min(CAPS.get(cid, 0.7) for cid in failures))
    return {
        "task_id": "cowork-evo/cowork-bench-course-enrollment-dashboard",
        "rubric_version": "cowork-bench-course-enrollment-dashboard-v0.1",
        "judge_version": "cowork-bench-course-enrollment-dashboard-v0.1",
        "criterion_results": [{"criterion_id": c["criterion_id"], "score": c["score"], "evidence": c["evidence"], "method": c["method"], "raw": c} for c in checks],
        "reward": round(reward, 6),
        "passed": not failures and reward >= 0.8,
        "hurdle_failures": failures,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    payload = evaluate(args.artifact_dir)
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{payload['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
