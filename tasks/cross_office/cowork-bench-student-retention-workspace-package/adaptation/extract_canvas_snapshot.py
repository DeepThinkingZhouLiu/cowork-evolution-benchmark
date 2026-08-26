#!/usr/bin/env python3
"""从 Cowork-Bench db/init.sql.gz 逐行投影课程 16/17 所需的 Canvas 源记录。"""

from __future__ import annotations

import argparse
import csv
import gzip
import re
from pathlib import Path


COURSE_IDS = {16, 17}
TABLES = {"assignments", "courses", "enrollments", "submissions"}
COPY_RE = re.compile(r'^COPY canvas\.(\w+) \((.+)\) FROM stdin;$')


def decode_copy(value: str) -> str:
    if value == r"\N":
        return ""
    return value.replace(r"\t", "\t").replace(r"\n", "\n").replace(r"\r", "\r").replace(r"\\", "\\")


def write_csv(path: Path, headers: list[str], rows: list[list]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\n")
        writer.writerow(headers)
        writer.writerows(rows)


def extract(sql_gz: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    section = None
    columns: dict[str, list[str]] = {}
    assignment_course: dict[int, int] = {}
    course_rows, enrollment_rows, assignment_rows, submission_rows = [], [], [], []
    with gzip.open(sql_gz, "rt", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.rstrip("\n")
            match = COPY_RE.match(line)
            if match:
                table = match.group(1)
                columns[table] = [value.strip().strip('"') for value in match.group(2).split(",")]
                section = table if table in TABLES else None
                continue
            if line == r"\.":
                section = None
                continue
            if section is None:
                continue
            values = line.split("\t")
            index = {name: position for position, name in enumerate(columns[section])}
            if section == "assignments":
                course_id = int(values[index["course_id"]])
                if course_id in COURSE_IDS:
                    assignment_id = int(values[index["id"]])
                    assignment_course[assignment_id] = course_id
                    assignment_rows.append([
                        assignment_id, course_id, decode_copy(values[index["name"]]),
                        decode_copy(values[index["points_possible"]]),
                    ])
            elif section == "courses":
                course_id = int(values[index["id"]])
                if course_id in COURSE_IDS:
                    course_rows.append([
                        course_id, decode_copy(values[index["name"]]),
                        decode_copy(values[index["course_code"]]), int(values[index["total_students"]]),
                    ])
            elif section == "enrollments":
                course_id = int(values[index["course_id"]])
                enrollment_type = decode_copy(values[index["type"]])
                if course_id in COURSE_IDS and "student" in enrollment_type.lower():
                    enrollment_rows.append([
                        int(values[index["id"]]), course_id, int(values[index["user_id"]]), enrollment_type,
                    ])
            elif section == "submissions":
                assignment_id = int(values[index["assignment_id"]])
                if assignment_id in assignment_course:
                    score = decode_copy(values[index["score"]])
                    submission_rows.append([
                        int(values[index["id"]]), assignment_id, int(values[index["user_id"]]), score,
                    ])
    if [len(course_rows), len(enrollment_rows), len(assignment_rows), len(submission_rows)] != [2, 4648, 26, 32424]:
        raise ValueError(
            f"源快照规模与审计值不一致: courses={len(course_rows)}, enrollments={len(enrollment_rows)}, "
            f"assignments={len(assignment_rows)}, submissions={len(submission_rows)}"
        )
    write_csv(output_dir / "canvas_courses.csv", ["course_id", "course_name", "course_code", "total_students"], course_rows)
    write_csv(output_dir / "canvas_enrollments.csv", ["enrollment_id", "course_id", "user_id", "type"], enrollment_rows)
    write_csv(output_dir / "canvas_assignments.csv", ["assignment_id", "course_id", "assignment_name", "points_possible"], assignment_rows)
    write_csv(output_dir / "canvas_submissions.csv", ["submission_id", "assignment_id", "user_id", "score"], submission_rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sql-gz", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    extract(args.sql_gz, args.output_dir)


if __name__ == "__main__":
    main()
