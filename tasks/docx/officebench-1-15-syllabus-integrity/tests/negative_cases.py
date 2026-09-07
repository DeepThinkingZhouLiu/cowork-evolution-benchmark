#!/usr/bin/env python3
"""构造并验证课程政策整合题的专属负例。"""

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
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
PYTHON = sys.executable
OUTPUT_NAME = "sample_syllabus_policy_review.docx"
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def run_judge(case: Path) -> float:
    result = case.parent / f"{case.name}-result.json"
    reward = case.parent / f"{case.name}-reward.txt"
    subprocess.run(
        [
            PYTHON,
            str(TASK_ROOT / "tests/judge.py"),
            "--artifact-dir",
            str(case),
            "--result",
            str(result),
            "--reward-file",
            str(reward),
        ],
        check=False,
        timeout=360,
    )
    return float(json.loads(result.read_text(encoding="utf-8"))["reward"])


def clone(source: Path, root: Path, name: str) -> Path:
    case = root / name
    case.mkdir()
    shutil.copy2(source, case / OUTPUT_NAME)
    return case


def rewrite_part(path: Path, part_name: str, mutator) -> None:
    temporary = path.with_suffix(".rewrite.docx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == part_name:
                root = etree.fromstring(payload)
                mutator(root)
                payload = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
            target.writestr(info, payload)
    temporary.replace(path)


def unwrap_policy_insertions(root) -> None:
    for insertion in root.xpath("//w:ins[@w:author='Course Policy Editor']", namespaces=W_NS):
        parent = insertion.getparent()
        index = parent.index(insertion)
        for child in list(insertion):
            parent.insert(index, child)
            index += 1
        parent.remove(insertion)


def move_policy_after_special_needs(root) -> None:
    body = root.find("w:body", namespaces=W_NS)
    tracked = [
        paragraph
        for paragraph in body.findall("w:p", namespaces=W_NS)
        if paragraph.xpath("./w:ins[@w:author='Course Policy Editor']", namespaces=W_NS)
    ]
    special = next(
        paragraph
        for paragraph in body.findall("w:p", namespaces=W_NS)
        if "".join(paragraph.xpath(".//w:t/text()", namespaces=W_NS)).startswith("SPECIAL NEEDS Accommodations")
    )
    for paragraph in tracked:
        body.remove(paragraph)
    position = body.index(special) + 1
    for paragraph in tracked:
        body.insert(position, paragraph)
        position += 1


def drop_anonymity_policy_paragraph(root) -> None:
    for paragraph in root.xpath("//w:body/w:p[w:ins[@w:author='Course Policy Editor']]", namespaces=W_NS):
        text = "".join(paragraph.xpath(".//w:t/text()", namespaces=W_NS))
        if text.startswith("Protect the anonymity"):
            paragraph.getparent().remove(paragraph)
            return


def remove_comment_anchors(root) -> None:
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


def replace_logo_payload(path: Path) -> None:
    tiny_jpeg = base64.b64decode(
        "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////"
        "2wBDAf//////////////////////////////////////////////////////////////////////////////////////wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/"
        "xAAUEAEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIQAxAAAAEf/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABBQJ//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPwF//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPwF//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQAGPwJ//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPyF//9oADAMBAAIAAwAAABAf/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAwEBPxB//8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAgBAgEBPxB//8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxB//9k="
    )
    temporary = path.with_suffix(".logo.docx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            payload = tiny_jpeg if info.filename == "word/media/image1.jpeg" else source.read(info.filename)
            target.writestr(info, payload)
    temporary.replace(path)


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="syllabus-policy-negative-") as temporary:
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
        source = fresh / OUTPUT_NAME
        rewards: dict[str, float] = {}

        missing = root / "missing_output"
        missing.mkdir()
        rewards[missing.name] = run_judge(missing)

        corrupt = root / "corrupt_docx"
        corrupt.mkdir()
        (corrupt / OUTPUT_NAME).write_bytes(b"not a docx")
        rewards[corrupt.name] = run_judge(corrupt)

        protected = clone(source, root, "protected_template_rewritten")
        document = Document(protected / OUTPUT_NAME)
        document.paragraphs[24].text = document.paragraphs[24].text.replace("socially just", "commercial")
        document.save(protected / OUTPUT_NAME)
        rewards[protected.name] = run_judge(protected)

        schedule = clone(source, root, "protected_schedule_damaged")
        document = Document(schedule / OUTPUT_NAME)
        document.tables[0].rows[2].cells[0].text = "Week X"
        document.save(schedule / OUTPUT_NAME)
        rewards[schedule.name] = run_judge(schedule)

        logo = clone(source, root, "protected_logo_replaced")
        replace_logo_payload(logo / OUTPUT_NAME)
        rewards[logo.name] = run_judge(logo)

        untracked = clone(source, root, "policy_insertions_untracked")
        rewrite_part(untracked / OUTPUT_NAME, "word/document.xml", unwrap_policy_insertions)
        rewards[untracked.name] = run_judge(untracked)

        wrong_anchor = clone(source, root, "policy_inserted_after_wrong_section")
        rewrite_part(wrong_anchor / OUTPUT_NAME, "word/document.xml", move_policy_after_special_needs)
        rewards[wrong_anchor.name] = run_judge(wrong_anchor)

        partial = clone(source, root, "anonymity_principle_missing")
        rewrite_part(partial / OUTPUT_NAME, "word/document.xml", drop_anonymity_policy_paragraph)
        rewards[partial.name] = run_judge(partial)

        unsupported = clone(source, root, "unsupported_irb_requirement")
        document = Document(unsupported / OUTPUT_NAME)
        scope = next(paragraph for paragraph in document.paragraphs if paragraph.text.startswith("The four policy components"))
        scope.add_run(" Institutional Review Board approval is mandatory before every class project.")
        document.save(unsupported / OUTPUT_NAME)
        rewards[unsupported.name] = run_judge(unsupported)

        wrong_matrix = clone(source, root, "anonymity_control_wrong")
        document = Document(wrong_matrix / OUTPUT_NAME)
        matrix = document.tables[-1]
        matrix.rows[6].cells[1].text = "Participant names may be disclosed after collection."
        matrix.rows[6].cells[2].text = "Publish all participant identities."
        document.save(wrong_matrix / OUTPUT_NAME)
        rewards[wrong_matrix.name] = run_judge(wrong_matrix)

        comments = clone(source, root, "scope_comments_removed")
        rewrite_part(comments / OUTPUT_NAME, "word/document.xml", remove_comment_anchors)
        rewards[comments.name] = run_judge(comments)

        medium = clone(source, root, "medium_missing_internal_links")
        rewrite_part(medium / OUTPUT_NAME, "word/document.xml", remove_internal_links)
        rewards[medium.name] = run_judge(medium)

        flattened = root / "flattened_picture_only"
        flattened.mkdir()
        pixel = root / "pixel.png"
        pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document()
        document.add_heading("Syllabus Policy Integration Review", level=1)
        document.add_picture(str(pixel), width=Inches(1))
        document.save(flattened / OUTPUT_NAME)
        rewards[flattened.name] = run_judge(flattened)

        assert rewards["missing_output"] == 0.0, rewards
        assert rewards["corrupt_docx"] == 0.0, rewards
        assert rewards["protected_template_rewritten"] <= 0.35, rewards
        assert rewards["protected_schedule_damaged"] <= 0.35, rewards
        assert rewards["protected_logo_replaced"] <= 0.35, rewards
        assert rewards["policy_insertions_untracked"] <= 0.62, rewards
        assert rewards["policy_inserted_after_wrong_section"] <= 0.62, rewards
        assert rewards["anonymity_principle_missing"] <= 0.62, rewards
        assert rewards["unsupported_irb_requirement"] <= 0.78, rewards
        assert rewards["anonymity_control_wrong"] <= 0.72, rewards
        assert rewards["scope_comments_removed"] <= 0.78, rewards
        assert 0.0 < rewards["medium_missing_internal_links"] < 1.0, rewards
        assert rewards["flattened_picture_only"] < 0.8, rewards
        print(json.dumps(rewards, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
