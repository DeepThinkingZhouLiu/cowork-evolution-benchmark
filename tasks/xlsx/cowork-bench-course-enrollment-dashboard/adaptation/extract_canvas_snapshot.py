#!/usr/bin/env python3
"""从固定 Cowork-Bench Canvas SQL 快照投影课程 16/17 的可见字段。

脚本只投影原始 COPY 行，不生成学生、成绩或作业记录。它保留原题评测
所需的 current_score、workflow_state 和 late 字段，使离线 XLSX 仍能按
原始查询口径重算。
"""

from __future__ import annotations

import argparse
import csv
import gzip
import json
import re
from pathlib import Path


COURSE_IDS = {16, 17}
COPY_RE = re.compile(r'^COPY canvas\.(\w+) \((.+)\) FROM stdin;$')


def decode(value: str) -> str:
    if value == r"\N":
        return ""
    return value.replace(r"\t", "\t").replace(r"\n", "\n").replace(r"\r", "\r").replace(r"\\", "\\")


def write_csv(path: Path, headers: list[str], rows: list[list[str | int]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.writer(stream, lineterminator="\n")
        writer.writerow(headers)
        writer.writerows(rows)


def extract(sql_gz: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    section: str | None = None
    columns: dict[str, list[str]] = {}
    assignment_course: dict[int, int] = {}
    courses: list[list[str | int]] = []
    enrollments: list[list[str | int]] = []
    assignments: list[list[str | int]] = []
    submissions: list[list[str | int]] = []
    with gzip.open(sql_gz, "rt", encoding="utf-8") as stream:
        for raw in stream:
            line = raw.rstrip("\n")
            match = COPY_RE.match(line)
            if match:
                section = match.group(1)
                columns[section] = [part.strip().strip('"') for part in match.group(2).split(",")]
                continue
            if line == r"\.":
                section = None
                continue
            if section not in {"courses", "enrollments", "assignments", "submissions"}:
                continue
            values = line.split("\t")
            index = {name: pos for pos, name in enumerate(columns[section])}
            if section == "courses":
                course_id = int(values[index["id"]])
                if course_id in COURSE_IDS:
                    courses.append([
                        course_id,
                        decode(values[index["name"]]),
                        decode(values[index["course_code"]]),
                        decode(values[index["workflow_state"]]),
                        int(values[index["total_students"]] or 0),
                    ])
            elif section == "enrollments":
                course_id = int(values[index["course_id"]])
                if course_id not in COURSE_IDS:
                    continue
                grades = decode(values[index["grades"]])
                current_score = ""
                if grades and grades != "{}":
                    try:
                        current_score = json.loads(grades).get("current_score", "")
                    except json.JSONDecodeError:
                        current_score = ""
                enrollments.append([
                    int(values[index["id"]]), course_id, int(values[index["user_id"]]),
                    decode(values[index["type"]]), decode(values[index["enrollment_state"]]),
                    current_score,
                ])
            elif section == "assignments":
                course_id = int(values[index["course_id"]])
                if course_id in COURSE_IDS:
                    assignment_id = int(values[index["id"]])
                    assignment_course[assignment_id] = course_id
                    assignments.append([
                        assignment_id, course_id, decode(values[index["name"]]),
                        decode(values[index["points_possible"]]),
                    ])
            elif section == "submissions":
                assignment_id = int(values[index["assignment_id"]])
                if assignment_id in assignment_course:
                    submissions.append([
                        int(values[index["id"]]), assignment_id,
                        int(values[index["user_id"]]), decode(values[index["score"]]),
                        decode(values[index["workflow_state"]]), decode(values[index["late"]]),
                        assignment_course[assignment_id],
                    ])

    courses.sort(key=lambda row: int(row[0]))
    enrollments.sort(key=lambda row: int(row[0]))
    assignments.sort(key=lambda row: int(row[0]))
    submissions.sort(key=lambda row: int(row[0]))
    # 4,648 student rows plus six teacher/TA rows are retained so that the
    # workbook can distinguish total enrollment from student enrollment.
    expected = (2, 4654, 26, 32424)
    actual = (len(courses), len(enrollments), len(assignments), len(submissions))
    if actual != expected:
        raise ValueError(f"源快照规模不符: actual={actual}, expected={expected}")
    write_csv(output_dir / "canvas_courses.csv", ["course_id", "course_name", "course_code", "workflow_state", "total_students"], courses)
    write_csv(output_dir / "canvas_enrollments.csv", ["enrollment_id", "course_id", "user_id", "type", "enrollment_state", "current_score"], enrollments)
    write_csv(output_dir / "canvas_assignments.csv", ["assignment_id", "course_id", "assignment_name", "points_possible"], assignments)
    write_csv(output_dir / "canvas_submissions.csv", ["submission_id", "assignment_id", "user_id", "score", "workflow_state", "late", "course_id"], submissions)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sql-gz", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    extract(args.sql_gz, args.output_dir)


if __name__ == "__main__":
    main()
