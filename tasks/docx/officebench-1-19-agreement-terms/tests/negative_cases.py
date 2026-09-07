#!/usr/bin/env python3
"""为 agreement-terms v0.3 独立构造定向负例与中等质量样例。"""

from __future__ import annotations

import argparse
import base64
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_NAME = "agreement_terms_audit.docx"
PYTHON = sys.executable
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def score(case_dir: Path) -> float:
    result = case_dir.parent / f"{case_dir.name}-result.json"
    reward = case_dir.parent / f"{case_dir.name}-reward.txt"
    subprocess.run(
        [
            PYTHON,
            str(TASK_ROOT / "tests/judge.py"),
            "--artifact-dir",
            str(case_dir),
            "--result",
            str(result),
            "--reward-file",
            str(reward),
        ],
        check=False,
        timeout=360,
    )
    return float(json.loads(result.read_text(encoding="utf-8"))["reward"])


def clone(golden: Path, root: Path, name: str) -> Path:
    case = root / name
    case.mkdir()
    shutil.copy2(golden, case / OUTPUT_NAME)
    return case


def replace_all(path: Path, before: str, after: str) -> None:
    document = Document(path)
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if before in run.text:
                run.text = run.text.replace(before, after)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if before in run.text:
                            run.text = run.text.replace(before, after)
    document.save(path)


def rewrite_document_xml(path: Path, mutator) -> None:
    temporary = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == "word/document.xml":
                root = etree.fromstring(payload)
                mutator(root)
                payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            target.writestr(info, payload)
    temporary.replace(path)


def remove_navigation(root) -> None:
    for node in root.xpath("//w:bookmarkStart | //w:bookmarkEnd", namespaces=W_NS):
        node.getparent().remove(node)
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=W_NS):
        parent = hyperlink.getparent()
        position = parent.index(hyperlink)
        for child in list(hyperlink):
            parent.insert(position, child)
            position += 1
        parent.remove(hyperlink)


def remove_matrix_row(root) -> None:
    matrix = root.xpath("//w:body/w:tbl[3]", namespaces=W_NS)[0]
    rows = matrix.xpath("./w:tr", namespaces=W_NS)
    matrix.remove(rows[9])


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="agreement-terms-negative-") as temporary:
        root = Path(temporary)
        fresh = root / "fresh"
        fresh.mkdir()
        subprocess.run(
            [
                PYTHON,
                str(TASK_ROOT / "solution/solve.py"),
                "--input-dir",
                str(TASK_ROOT / "data/input_files"),
                "--output-dir",
                str(fresh),
            ],
            check=True,
            timeout=180,
        )
        golden = fresh / OUTPUT_NAME
        cases: dict[str, float] = {}

        missing = root / "missing_output"
        missing.mkdir()
        cases[missing.name] = score(missing)

        corrupt = root / "corrupt_docx"
        corrupt.mkdir()
        (corrupt / OUTPUT_NAME).write_bytes(b"not a docx")
        cases[corrupt.name] = score(corrupt)

        extra = clone(golden, root, "extra_artifact")
        (extra / "answer.txt").write_text("extra", encoding="utf-8")
        cases[extra.name] = score(extra)

        swapped = clone(golden, root, "party_roles_swapped")
        replace_all(swapped / OUTPUT_NAME, "Provider = Acme Cloud Services, Inc.; Customer = Wayne Enterprises, Inc.", "Provider = Wayne Enterprises, Inc.; Customer = Acme Cloud Services, Inc.")
        cases[swapped.name] = score(swapped)

        order = clone(golden, root, "order_precedence_reversed")
        replace_all(order / OUTPUT_NAME, "Order Form 对该 Agreement 的不一致之处优先于 Framework Terms", "Framework Terms 在不一致之处优先于 Order Form")
        cases[order.name] = score(order)

        dpa = clone(golden, root, "dpa_precedence_truncated")
        replace_all(dpa / OUTPUT_NAME, "EEA SCCs/UK Addendum > DPA > Agreement", "DPA > Agreement")
        cases[dpa.name] = score(dpa)

        incident = clone(golden, root, "incident_changed_to_business_hours")
        replace_all(incident / OUTPUT_NAME, "72 小时", "72 个工作小时")
        cases[incident.name] = score(incident)

        subprocessor = clone(golden, root, "subprocessor_windows_swapped")
        replace_all(subprocessor / OUTPUT_NAME, "至少提前 10 个工作日通知；Customer 在通知后 30 日内可异议", "至少提前 30 日通知；Customer 在通知后 10 个工作日内可异议")
        cases[subprocessor.name] = score(subprocessor)

        deletion = clone(golden, root, "deletion_mechanisms_collapsed")
        replace_all(deletion / OUTPUT_NAME, "DPA 到期后按 Customer 指示返还或删除；法律要求保留时例外", "DPA 到期后无条件在 60 日内删除")
        cases[deletion.name] = score(deletion)

        cap = clone(golden, root, "fabricated_liability_multiplier")
        replace_all(cap / OUTPUT_NAME, "[Fill in a number]", "2x the Fees")
        cases[cap.name] = score(cap)

        law = clone(golden, root, "fabricated_governing_law")
        replace_all(law / OUTPUT_NAME, "州、国家与法院所在地均未填", "laws of Delaware; courts in Delaware")
        cases[law.name] = score(law)

        missing_row = clone(golden, root, "matrix_row_missing")
        rewrite_document_xml(missing_row / OUTPUT_NAME, remove_matrix_row)
        cases[missing_row.name] = score(missing_row)

        flattened = root / "flattened_screenshot_only"
        flattened.mkdir()
        pixel = root / "pixel.png"
        pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document()
        document.add_heading("Cloud Agreement Terms & Readiness Audit", 0)
        document.add_picture(str(pixel), width=Inches(6))
        document.save(flattened / OUTPUT_NAME)
        cases[flattened.name] = score(flattened)

        section = clone(golden, root, "landscape_matrix_destroyed")
        document = Document(section / OUTPUT_NAME)
        document.sections[1].orientation = WD_ORIENT.PORTRAIT
        document.sections[1].page_width, document.sections[1].page_height = document.sections[1].page_height, document.sections[1].page_width
        document.save(section / OUTPUT_NAME)
        cases[section.name] = score(section)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite_document_xml(medium / OUTPUT_NAME, remove_navigation)
        cases[medium.name] = score(medium)

        assert cases["missing_output"] == 0.0, cases
        assert cases["corrupt_docx"] == 0.0, cases
        assert cases["extra_artifact"] == 0.0, cases
        assert 0 < cases["party_roles_swapped"] <= 0.35, cases
        assert 0 < cases["order_precedence_reversed"] <= 0.55, cases
        assert 0 < cases["dpa_precedence_truncated"] <= 0.55, cases
        assert 0 < cases["incident_changed_to_business_hours"] < 0.8, cases
        assert 0 < cases["subprocessor_windows_swapped"] < 0.8, cases
        assert 0 < cases["deletion_mechanisms_collapsed"] < 0.8, cases
        assert 0 < cases["fabricated_liability_multiplier"] <= 0.75, cases
        assert 0 < cases["fabricated_governing_law"] <= 0.75, cases
        assert 0 < cases["matrix_row_missing"] <= 0.65, cases
        assert cases["flattened_screenshot_only"] <= 0.35, cases
        assert 0 < cases["landscape_matrix_destroyed"] < 1.0, cases
        assert 0 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--retain-dir", type=Path)
    arguments = parser.parse_args()
    run(arguments.retain_dir)
