#!/usr/bin/env python3
"""Project only C095-relevant Canvas rows from Cowork-Bench's frozen SQL dump."""

from __future__ import annotations

import argparse
import csv
import gzip
import hashlib
import html
import json
import re
from pathlib import Path


SUBJECT_AREAS = {
    "AAA": "Analytics",
    "BBB": "Biochemistry",
    "CCC": "Computing",
    "DDD": "Design",
    "EEE": "Economics",
    "FFF": "Finance",
    "GGG": "Governance",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def unescape_copy(value: str):
    if value == r"\N":
        return None
    return (
        value.replace(r"\t", "\t")
        .replace(r"\n", "\n")
        .replace(r"\\", "\\")
    )


def copy_rows(sql: str, table: str) -> list[list[str | None]]:
    marker = f"COPY canvas.{table} "
    start = sql.index(marker)
    end = sql.index("\n\\.\n", start)
    rows = []
    for line in sql[start:end].splitlines()[1:]:
        rows.append([unescape_copy(value) for value in line.split("\t")])
    return rows


def write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def course_subject(code: str, name: str) -> str:
    prefix = code.split("-", 1)[0]
    if prefix not in SUBJECT_AREAS:
        raise ValueError(f"Unsupported course prefix {prefix!r} for {name!r}")
    return SUBJECT_AREAS[prefix]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sql-gz", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    sql = gzip.open(args.sql_gz, "rt", encoding="utf-8").read()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    user_rows = copy_rows(sql, "users")
    course_rows = copy_rows(sql, "courses")
    enrollment_rows = copy_rows(sql, "enrollments")
    assignment_rows = copy_rows(sql, "assignments")

    users = {int(row[0]): row for row in user_rows}
    courses = {int(row[0]): row for row in course_rows}
    teacher_rows = [row for row in enrollment_rows if row[3] == "TeacherEnrollment"]
    teacher_ids = {int(row[2]) for row in teacher_rows}
    student_rows = [row for row in enrollment_rows if row[3] == "StudentEnrollment"]

    course_out = []
    for row in sorted(course_rows, key=lambda item: int(item[0])):
        code = str(row[3])
        # The subject area is the stable prefix of the visible Canvas course code.
        course_out.append(
            {
                "Course_ID": row[0],
                "Course_Code": code,
                "Course_Name": html.unescape(str(row[2] or "")),
                "Subject_Area": course_subject(code, str(row[2] or "")) if code.split("-", 1)[0] in SUBJECT_AREAS else "Other",
                "Student_Count_Source": row[15] or "0",
                "Workflow_State": row[4] or "",
            }
        )

    teacher_out = []
    for row in sorted(teacher_rows, key=lambda item: (str(users[int(item[2])][1]), int(item[1]))):
        user = users[int(row[2])]
        course = courses[int(row[1])]
        teacher_out.append(
            {
                "Teacher_ID": row[2],
                "Instructor": user[1],
                "Login_ID": user[4],
                "Course_ID": row[1],
                "Course_Code": course[3],
            }
        )

    student_out = [
        {"Course_ID": row[1], "Student_ID": row[2], "Enrollment_State": row[4] or ""}
        for row in sorted(student_rows, key=lambda item: (int(item[1]), int(item[2])))
    ]
    assignment_out = [
        {"Assignment_ID": row[0], "Course_ID": row[1], "Assignment_Name": row[3] or ""}
        for row in sorted(assignment_rows, key=lambda item: int(item[0]))
    ]

    write_csv(args.output_dir / "canvas_courses.csv", list(course_out[0]), course_out)
    write_csv(args.output_dir / "canvas_teacher_enrollments.csv", list(teacher_out[0]), teacher_out)
    write_csv(args.output_dir / "canvas_student_enrollments.csv", list(student_out[0]), student_out)
    write_csv(args.output_dir / "canvas_assignments.csv", list(assignment_out[0]), assignment_out)

    manifest = {
        "source_sql_sha256": sha256(args.sql_gz),
        "source_path": "db/init.sql.gz",
        "teacher_count": len(teacher_ids),
        "teacher_enrollment_rows": len(teacher_out),
        "course_count": len(course_out),
        "student_enrollment_rows": len(student_out),
        "assignment_rows": len(assignment_out),
        "projection": "All TeacherEnrollment rows, all course rows, all StudentEnrollment rows, and all assignment rows from the fixed Canvas snapshot; no unrelated user fields retained.",
    }
    (args.output_dir / "canvas_snapshot_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
