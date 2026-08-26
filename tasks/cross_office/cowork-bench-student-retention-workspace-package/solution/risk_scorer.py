#!/usr/bin/env python3
"""Canvas CSV 快照的可回放学生风险评分器。"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path


COURSE_IDS = [16, 17]
RISK_ORDER = ["High", "Medium", "Low"]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def parse_thresholds(model: dict) -> tuple[float, float]:
    classification = model["risk_classification"]
    high_match = re.search(r"-?\d+(?:\.\d+)?", classification["high_risk"]["threshold"])
    medium_values = [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", classification["medium_risk"]["threshold"])]
    if not high_match or len(medium_values) < 2:
        raise ValueError("无法从 scoring_model.json 解析风险阈值")
    high_cutoff = float(high_match.group(0))
    low_cutoff = round(medium_values[1] + 0.01, 2)
    if medium_values[0] != high_cutoff:
        raise ValueError("高/中风险阈值不连续")
    return high_cutoff, low_cutoff


def classify(value: float, high_cutoff: float, low_cutoff: float) -> str:
    if value < high_cutoff:
        return "High"
    if value < low_cutoff:
        return "Medium"
    return "Low"


def score(courses_path: Path, enrollments_path: Path, assignments_path: Path, submissions_path: Path, model_path: Path) -> dict:
    courses = {int(row["course_id"]): row for row in read_csv(courses_path) if int(row["course_id"]) in COURSE_IDS}
    if sorted(courses) != COURSE_IDS:
        raise ValueError("课程快照必须恰好包含课程 16/17")
    enrollments = read_csv(enrollments_path)
    enrollment_counts = Counter(int(row["course_id"]) for row in enrollments if int(row["course_id"]) in COURSE_IDS and "student" in row["type"].lower())
    assignments = read_csv(assignments_path)
    assignment_course = {int(row["assignment_id"]): int(row["course_id"]) for row in assignments if int(row["course_id"]) in COURSE_IDS}
    if len(assignment_course) != 26:
        raise ValueError(f"预期 26 个作业，实际 {len(assignment_course)}")
    model = json.loads(model_path.read_text(encoding="utf-8"))
    high_cutoff, low_cutoff = parse_thresholds(model)

    student_scores: dict[tuple[int, int], list[float]] = defaultdict(list)
    course_scores: dict[int, list[float]] = defaultdict(list)
    for row in read_csv(submissions_path):
        assignment_id = int(row["assignment_id"])
        if assignment_id not in assignment_course or row["score"].strip() == "":
            continue
        course_id = assignment_course[assignment_id]
        value = float(row["score"])
        student_scores[(course_id, int(row["user_id"]))].append(value)
        course_scores[course_id].append(value)

    per_course_risks: dict[int, Counter] = {course_id: Counter() for course_id in COURSE_IDS}
    for (course_id, _user_id), values in student_scores.items():
        per_course_risks[course_id][classify(sum(values) / len(values), high_cutoff, low_cutoff)] += 1

    overview, at_risk, tracker = [], [], []
    for course_id in COURSE_IDS:
        risks = per_course_risks[course_id]
        graded_students = sum(risks.values())
        avg_score = round(sum(course_scores[course_id]) / len(course_scores[course_id]), 2)
        pass_rate = round(100.0 * (risks["Medium"] + risks["Low"]) / graded_students, 2)
        predominant = min(RISK_ORDER, key=lambda level: (-risks[level], RISK_ORDER.index(level)))
        overview.append({
            "course_id": course_id,
            "course_name": courses[course_id]["course_name"],
            "enrollment_count": enrollment_counts[course_id],
            "graded_student_count": graded_students,
            "graded_submission_count": len(course_scores[course_id]),
            "avg_score": avg_score,
            "pass_rate": pass_rate,
        })
        at_risk.append({
            "course_name": courses[course_id]["course_name"],
            "high_risk_count": risks["High"],
            "medium_risk_count": risks["Medium"],
            "low_risk_count": risks["Low"],
        })
        tracker.append({
            "Student Course": courses[course_id]["course_name"],
            "Risk Level": predominant,
            "Student Count": graded_students,
            "Average Score": avg_score,
            "Pass Rate": pass_rate,
        })

    overall = Counter()
    for risks in per_course_risks.values():
        overall.update(risks)
    denominator = sum(overall.values())
    distribution = [
        {"risk_level": level, "student_count": overall[level], "pct": round(100.0 * overall[level] / denominator, 2)}
        for level in RISK_ORDER
    ]
    actions = model["intervention_actions"]
    intervention = [
        {"risk_level": level, **actions[level.lower()]}
        for level in RISK_ORDER
    ]
    return {
        "course_overview": overview,
        "risk_distribution": distribution,
        "at_risk_students": at_risk,
        "intervention_plan": intervention,
        "student_risk_tracker": tracker,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--courses", type=Path, required=True)
    parser.add_argument("--enrollments", type=Path, required=True)
    parser.add_argument("--assignments", type=Path, required=True)
    parser.add_argument("--submissions", type=Path, required=True)
    parser.add_argument("--model", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = score(args.courses, args.enrollments, args.assignments, args.submissions, args.model)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
