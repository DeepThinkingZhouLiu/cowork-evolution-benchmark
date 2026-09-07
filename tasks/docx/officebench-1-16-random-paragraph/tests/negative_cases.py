#!/usr/bin/env python3
"""构造并验证本题专属的定向负例与中等样例。"""

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
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_NAME = "cafe_profile_editorial_review.docx"
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


def accept_all_revisions(root) -> None:
    for deleted in root.xpath("//w:del", namespaces=W_NS):
        deleted.getparent().remove(deleted)
    for inserted in root.xpath("//w:ins", namespaces=W_NS):
        parent = inserted.getparent()
        index = parent.index(inserted)
        for child in list(inserted):
            if child.tag in {qn("w:commentRangeStart"), qn("w:commentRangeEnd")}:
                continue
            if child.tag == qn("w:r") and child.find(qn("w:commentReference")) is not None:
                continue
            parent.insert(index, child)
            index += 1
        parent.remove(inserted)


def remove_s5_revision(root) -> None:
    for paragraph in root.xpath("//w:p[w:ins]", namespaces=W_NS):
        label = "".join(paragraph.xpath("./w:r/w:t/text()", namespaces=W_NS))
        if label.startswith("S5 |"):
            for node in paragraph.xpath("./w:del | ./w:ins", namespaces=W_NS):
                paragraph.remove(node)


def wrong_s3_approved_text(root) -> None:
    for paragraph in root.xpath("//w:p[w:ins]", namespaces=W_NS):
        label = "".join(paragraph.xpath("./w:r/w:t/text()", namespaces=W_NS))
        if not label.startswith("S3 |"):
            continue
        nodes = paragraph.xpath("./w:ins//w:t", namespaces=W_NS)
        nodes[-1].text = (nodes[-1].text or "") + " Croissants cost $5."


def remove_comment_ranges(root) -> None:
    for node in root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=W_NS):
        node.getparent().remove(node)


def remove_navigation(root) -> None:
    for node in root.xpath("//w:bookmarkStart | //w:bookmarkEnd", namespaces=W_NS):
        node.getparent().remove(node)
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=W_NS):
        parent = hyperlink.getparent()
        index = parent.index(hyperlink)
        for child in list(hyperlink):
            parent.insert(index, child)
            index += 1
        parent.remove(hyperlink)


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="cafe-profile-negative-") as temporary:
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
            timeout=120,
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

        extra_artifact = clone(golden, root, "extra_artifact")
        (extra_artifact / "answer.txt").write_text("extra", encoding="utf-8")
        cases[extra_artifact.name] = score(extra_artifact)

        protected = clone(golden, root, "controlled_source_changed")
        document = Document(protected / OUTPUT_NAME)
        document.tables[1].cell(0, 0).text = document.tables[1].cell(0, 0).text.replace("small café", "large café")
        document.save(protected / OUTPUT_NAME)
        cases[protected.name] = score(protected)

        accepted = clone(golden, root, "revisions_accepted")
        rewrite_document_xml(accepted / OUTPUT_NAME, accept_all_revisions)
        cases[accepted.name] = score(accepted)

        partial = clone(golden, root, "s5_revision_missing")
        rewrite_document_xml(partial / OUTPUT_NAME, remove_s5_revision)
        cases[partial.name] = score(partial)

        wrong_copy = clone(golden, root, "wrong_s3_approved_copy")
        rewrite_document_xml(wrong_copy / OUTPUT_NAME, wrong_s3_approved_text)
        cases[wrong_copy.name] = score(wrong_copy)

        wrong_count = clone(golden, root, "wrong_word_count_and_delta")
        document = Document(wrong_count / OUTPUT_NAME)
        document.tables[2].rows[3].cells[3].text = "14"
        document.tables[2].rows[3].cells[4].text = "-7"
        document.save(wrong_count / OUTPUT_NAME)
        cases[wrong_count.name] = score(wrong_count)

        comments = clone(golden, root, "comments_removed")
        rewrite_document_xml(comments / OUTPUT_NAME, remove_comment_ranges)
        cases[comments.name] = score(comments)

        unsupported = clone(golden, root, "unsupported_price_added")
        document = Document(unsupported / OUTPUT_NAME)
        document.add_paragraph("Espresso costs $5 and the café is open every day.")
        document.save(unsupported / OUTPUT_NAME)
        cases[unsupported.name] = score(unsupported)

        flattened = root / "flattened_screenshot_only"
        flattened.mkdir()
        image = root / "pixel.png"
        image.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document()
        document.add_heading("Café Profile Editorial Review", 0)
        document.add_picture(str(image), width=Inches(1))
        document.save(flattened / OUTPUT_NAME)
        cases[flattened.name] = score(flattened)

        wrong_section = clone(golden, root, "landscape_section_destroyed")
        document = Document(wrong_section / OUTPUT_NAME)
        document.sections[1].orientation = WD_ORIENT.PORTRAIT
        document.sections[1].page_width, document.sections[1].page_height = document.sections[1].page_height, document.sections[1].page_width
        document.save(wrong_section / OUTPUT_NAME)
        cases[wrong_section.name] = score(wrong_section)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite_document_xml(medium / OUTPUT_NAME, remove_navigation)
        cases[medium.name] = score(medium)

        assert cases["missing_output"] == 0.0, cases
        assert cases["corrupt_docx"] == 0.0, cases
        assert cases["extra_artifact"] == 0.0, cases
        assert cases["controlled_source_changed"] <= 0.35, cases
        assert cases["revisions_accepted"] <= 0.58, cases
        assert 0.0 < cases["s5_revision_missing"] <= 0.58, cases
        assert cases["wrong_s3_approved_copy"] <= 0.58, cases
        assert 0.0 < cases["wrong_word_count_and_delta"] <= 0.70, cases
        assert cases["comments_removed"] <= 0.72, cases
        assert cases["unsupported_price_added"] <= 0.76, cases
        assert cases["flattened_screenshot_only"] <= 0.35, cases
        assert 0.0 < cases["landscape_section_destroyed"] < 1.0, cases
        assert 0.0 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--retain-dir", type=Path)
    arguments = parser.parse_args()
    run(arguments.retain_dir)
