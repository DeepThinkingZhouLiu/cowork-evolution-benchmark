#!/usr/bin/env python3
"""任务级确定性 Judge；语义与视觉质量将在正式版叠加独立 Judge。"""

from __future__ import annotations

import argparse
import json
import math
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


def ratio(checks):
    return 1.0 if not checks else sum(bool(item) for item in checks) / len(checks)


def inspect_docx(path):
    doc = Document(path)
    text = "\n".join([p.text for p in doc.paragraphs] + [c.text for t in doc.tables for r in t.rows for c in r.cells])
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    root = ElementTree.fromstring(xml)
    xml_text = "".join(root.itertext())
    return {
        "text": text + "\n" + xml_text,
        "tables": len(doc.tables),
        "headings": sum(p.style and p.style.name.startswith("Heading") for p in doc.paragraphs),
        "styled": sum(bool(p.style and p.style.name != "Normal") for p in doc.paragraphs),
        "redlines": xml.count(b"<w:ins") + xml.count(b"<w:del"),
    }


def inspect_xlsx(path):
    wb = load_workbook(path, data_only=False)
    text = []
    charts = 0
    cells = {}
    for ws in wb.worksheets:
        charts += len(ws._charts)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    text.append(str(cell.value))
                    cells[f"{ws.title}!{cell.coordinate}"] = cell.value
    result = {"text": "\n".join(text), "sheets": wb.sheetnames, "charts": charts, "cells": cells}
    wb.close()
    return result


def inspect_pptx(path):
    prs = Presentation(path)
    text = []
    shape_count = 0
    for slide in prs.slides:
        shape_count += len(slide.shapes)
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                text.append(shape.text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    text.extend(cell.text for cell in row.cells)
    return {"text": "\n".join(text), "slides": len(prs.slides), "shapes": shape_count}


def equal_value(actual, expected):
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        try:
            return math.isclose(float(actual), float(expected), rel_tol=1e-6, abs_tol=1e-6)
        except (TypeError, ValueError):
            return False
    return str(actual) == str(expected)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("artifact_dir", type=Path)
    parser.add_argument("--expectations", type=Path, default=Path(__file__).with_name("expectations.json"))
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    expected = json.loads(args.expectations.read_text(encoding="utf-8"))
    inspected, open_checks = {}, []
    for filename in expected["required_files"]:
        path = args.artifact_dir / filename
        if not path.is_file():
            open_checks.append(False)
            continue
        try:
            suffix = path.suffix.lower()
            if suffix == ".docx":
                inspected[filename] = inspect_docx(path)
            elif suffix == ".xlsx":
                inspected[filename] = inspect_xlsx(path)
            elif suffix == ".pptx":
                inspected[filename] = inspect_pptx(path)
            else:
                inspected[filename] = {"text": path.read_text(encoding="utf-8")}
            open_checks.append(True)
        except Exception:
            open_checks.append(False)

    content_checks, exact_checks, structure_checks, presentation_checks = [], [], [], []
    preservation_checks, penalty_checks = [], []
    for filename, file_expected in expected.get("files", {}).items():
        info = inspected.get(filename, {"text": ""})
        text = info.get("text", "")
        content_checks.extend(term in text for term in file_expected.get("contains", []))
        exact_checks.extend(term in text for term in file_expected.get("exact_terms", []))
        preservation_checks.extend(term in text for term in file_expected.get("preserve_terms", []))
        penalty_checks.extend(term not in text for term in file_expected.get("excludes", []))
        if "min_tables" in file_expected:
            structure_checks.append(info.get("tables", 0) >= file_expected["min_tables"])
        if "min_headings" in file_expected:
            structure_checks.append(info.get("headings", 0) >= file_expected["min_headings"])
        if "min_redlines" in file_expected:
            structure_checks.append(info.get("redlines", 0) >= file_expected["min_redlines"])
        if "sheet_names" in file_expected:
            structure_checks.append(all(name in info.get("sheets", []) for name in file_expected["sheet_names"]))
        if "min_slides" in file_expected:
            structure_checks.append(info.get("slides", 0) >= file_expected["min_slides"])
        if "max_slides" in file_expected:
            structure_checks.append(info.get("slides", 0) <= file_expected["max_slides"])
        if "min_charts" in file_expected:
            presentation_checks.append(info.get("charts", 0) >= file_expected["min_charts"])
        if "min_shapes" in file_expected:
            presentation_checks.append(info.get("shapes", 0) >= file_expected["min_shapes"])
        if "min_styled" in file_expected:
            presentation_checks.append(info.get("styled", 0) >= file_expected["min_styled"])
        for coordinate, value in file_expected.get("cells", {}).items():
            exact_checks.append(equal_value(info.get("cells", {}).get(coordinate), value))
        for coordinate, value in file_expected.get("formulas", {}).items():
            exact_checks.append(info.get("cells", {}).get(coordinate) == value)

    consistency_checks = []
    for rule in expected.get("shared_terms", []):
        consistency_checks.append(all(rule["term"] in inspected.get(name, {}).get("text", "") for name in rule["files"]))
    consistency_checks.extend(preservation_checks)

    scores = {
        "R001": ratio(open_checks),
        "R002": ratio(content_checks),
        "R003": ratio(exact_checks),
        "R004": ratio(structure_checks),
        "R005": ratio(presentation_checks),
        "R006": ratio(consistency_checks),
        "R007": ratio(penalty_checks),
    }
    positive_weights = {"R001": 2, "R002": 4, "R003": 4, "R004": 2, "R005": 2, "R006": 2}
    reward = sum(scores[key] * weight for key, weight in positive_weights.items()) / sum(positive_weights.values())
    reward = max(0.0, reward - (1.0 - scores["R007"]) * 3 / sum(positive_weights.values()))
    if scores["R001"] < 1:
        reward = 0.0
    passed = reward >= expected.get("pass_threshold", 0.7)
    result = {
        "task_id": expected["task_id"],
        "judge_version": "cowork-evo-deterministic-v0.1",
        "reward": round(reward, 6),
        "passed": passed,
        "criterion_results": [
            {"criterion_id": key, "score": round(value, 6), "method": "deterministic"}
            for key, value in scores.items()
        ],
    }
    output = json.dumps(result, ensure_ascii=False, indent=2)
    print(output)
    if args.output:
        args.output.write_text(output + "\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
