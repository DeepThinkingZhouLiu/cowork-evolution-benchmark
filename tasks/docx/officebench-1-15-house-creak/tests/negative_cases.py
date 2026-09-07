#!/usr/bin/env python3
"""构造并验证本题特有的负例与中等样例。"""

from __future__ import annotations

import base64
import json
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
PYTHON = sys.executable
OUTPUT_NAME = "house_creak_editorial_review.docx"
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


def clone_golden(golden: Path, root: Path, name: str) -> Path:
    case = root / name
    case.mkdir()
    shutil.copy2(golden, case / OUTPUT_NAME)
    return case


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


def remove_comments(root) -> None:
    for node in root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=W_NS):
        node.getparent().remove(node)


def remove_internal_links(root) -> None:
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=W_NS):
        parent = hyperlink.getparent()
        index = parent.index(hyperlink)
        for child in list(hyperlink):
            parent.insert(index, child)
            index += 1
        parent.remove(hyperlink)


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="house-creak-negative-") as temp:
        root = Path(temp)
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
            timeout=120,
        )
        golden = fresh / OUTPUT_NAME

        cases: dict[str, float] = {}

        missing = root / "missing_output"
        missing.mkdir()
        cases["missing_output"] = score(missing)

        corrupt = root / "corrupt_docx"
        corrupt.mkdir()
        (corrupt / OUTPUT_NAME).write_bytes(b"not a docx")
        cases["corrupt_docx"] = score(corrupt)

        protected = clone_golden(golden, root, "protected_story_changed")
        document = Document(protected / OUTPUT_NAME)
        document.paragraphs[2].text = document.paragraphs[2].text.replace("interest", "alarm")
        document.save(protected / OUTPUT_NAME)
        cases["protected_story_changed"] = score(protected)

        wrong_core = clone_golden(golden, root, "wrong_chronology_fact")
        document = Document(wrong_core / OUTPUT_NAME)
        document.tables[0].rows[6].cells[1].text = "The neighbour said the house had passed inspection."
        document.tables[0].rows[6].cells[3].text = "The family learns that the building was safe."
        document.save(wrong_core / OUTPUT_NAME)
        cases["wrong_chronology_fact"] = score(wrong_core)

        partial = clone_golden(golden, root, "partial_character_matrix")
        document = Document(partial / OUTPUT_NAME)
        table = document.tables[1]
        table._tbl.remove(table.rows[-1]._tr)
        document.save(partial / OUTPUT_NAME)
        cases["partial_character_matrix"] = score(partial)

        extra = clone_golden(golden, root, "unsupported_extra_event")
        document = Document(extra / OUTPUT_NAME)
        cells = document.tables[0].add_row().cells
        for cell, value in zip(cells, ["7", "Father warned the family before the move.", "No risk", "The disaster was expected.", "Unsupported"]):
            cell.text = value
        document.save(extra / OUTPUT_NAME)
        cases["unsupported_extra_event"] = score(extra)

        flattened = root / "flattened_screenshot_only"
        flattened.mkdir()
        png = root / "pixel.png"
        png.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document()
        document.add_paragraph("The House that Creaked editorial review")
        document.add_picture(str(png), width=Inches(1))
        document.save(flattened / OUTPUT_NAME)
        cases["flattened_screenshot_only"] = score(flattened)

        comments = clone_golden(golden, root, "comments_removed")
        rewrite_document_xml(comments / OUTPUT_NAME, remove_comments)
        cases["comments_removed"] = score(comments)

        medium = clone_golden(golden, root, "medium_missing_navigation")
        rewrite_document_xml(medium / OUTPUT_NAME, remove_internal_links)
        cases["medium_missing_navigation"] = score(medium)

        assert cases["missing_output"] == 0.0, cases
        assert cases["corrupt_docx"] == 0.0, cases
        assert cases["protected_story_changed"] <= 0.35, cases
        assert cases["wrong_chronology_fact"] <= 0.72, cases
        assert 0.0 < cases["partial_character_matrix"] <= 0.72, cases
        assert cases["unsupported_extra_event"] <= 0.72, cases
        assert cases["flattened_screenshot_only"] == 0.0, cases
        assert cases["comments_removed"] <= 0.72, cases
        assert 0.0 < cases["medium_missing_navigation"] < 1.0, cases
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
