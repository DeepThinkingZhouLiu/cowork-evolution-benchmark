#!/usr/bin/env python3
"""构造并校准本题专属的定向负例与中等样例。"""

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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_NAME = "coordination_meeting_minutes_audit.docx"
PYTHON = sys.executable
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def score(case_dir: Path) -> float:
    result = case_dir.parent / f"{case_dir.name}-result.json"
    reward = case_dir.parent / f"{case_dir.name}-reward.txt"
    subprocess.run(
        [
            PYTHON, str(TASK_ROOT / "tests/judge.py"), "--artifact-dir", str(case_dir),
            "--result", str(result), "--reward-file", str(reward),
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


def rewrite_docx(path: Path, mutator) -> None:
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


def remove_named_transcript_row(root) -> None:
    for row in root.xpath("//w:body/w:tbl[1]/w:tr", namespaces=W_NS):
        text = "".join(row.xpath(".//w:t/text()", namespaces=W_NS))
        if "许承柏领学：提升综合服务质量的闭环管理" in text:
            row.getparent().remove(row)
            return


def accept_revision(root) -> None:
    for deleted in root.xpath("//w:body/w:tbl[1]//w:del", namespaces=W_NS):
        deleted.getparent().remove(deleted)
    for inserted in root.xpath("//w:body/w:tbl[1]//w:ins", namespaces=W_NS):
        parent = inserted.getparent()
        position = parent.index(inserted)
        for child in list(inserted):
            parent.insert(position, child)
            position += 1
        parent.remove(inserted)


def change_revision_count(root) -> None:
    nodes = root.xpath("//w:body/w:tbl[1]//w:ins//w:t[text()='5']", namespaces=W_NS)
    if nodes:
        nodes[0].text = "6"


def remove_comment_anchor(root) -> None:
    for node in root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=W_NS):
        node.getparent().remove(node)


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


def make_border_wrong(root) -> None:
    bottom = root.xpath("//w:body/w:tbl[1]/w:tr[1]/w:tc/w:tcPr/w:tcBorders/w:bottom", namespaces=W_NS)[0]
    bottom.set(qn("w:sz"), "24")
    bottom.set(qn("w:color"), "FF0000")


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="coordination-minutes-negative-") as temporary:
        root = Path(temporary)
        fresh = root / "fresh"
        fresh.mkdir()
        subprocess.run(
            [
                PYTHON, str(TASK_ROOT / "solution/solve.py"), "--input-dir", str(TASK_ROOT / "data/input_files"),
                "--output-dir", str(fresh),
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

        header = clone(golden, root, "protected_location_changed")
        document = Document(header / OUTPUT_NAME)
        target = next(p for p in document.paragraphs if "综合楼二层会议室" in p.text)
        target.text = target.text.replace("综合楼二层会议室", "综合楼一层会议室")
        document.save(header / OUTPUT_NAME)
        cases[header.name] = score(header)

        missing_line = clone(golden, root, "transcript_source_line_missing")
        rewrite_docx(missing_line / OUTPUT_NAME, remove_named_transcript_row)
        cases[missing_line.name] = score(missing_line)

        accepted = clone(golden, root, "count_revision_accepted")
        rewrite_docx(accepted / OUTPUT_NAME, accept_revision)
        cases[accepted.name] = score(accepted)

        wrong_count = clone(golden, root, "count_changed_to_six")
        rewrite_docx(wrong_count / OUTPUT_NAME, change_revision_count)
        cases[wrong_count.name] = score(wrong_count)

        no_comment = clone(golden, root, "count_comment_anchor_removed")
        rewrite_docx(no_comment / OUTPUT_NAME, remove_comment_anchor)
        cases[no_comment.name] = score(no_comment)

        blank_row = clone(golden, root, "blank_transcript_row_added")
        document = Document(blank_row / OUTPUT_NAME)
        document.tables[0].add_row()
        document.save(blank_row / OUTPUT_NAME)
        cases[blank_row.name] = score(blank_row)

        border = clone(golden, root, "first_rule_thick_red")
        rewrite_docx(border / OUTPUT_NAME, make_border_wrong)
        cases[border.name] = score(border)

        wrong_owner = clone(golden, root, "deviation_owner_swapped")
        document = Document(wrong_owner / OUTPUT_NAME)
        document.tables[2].cell(1, 4).text = "陆清予"
        document.save(wrong_owner / OUTPUT_NAME)
        cases[wrong_owner.name] = score(wrong_owner)

        wrong_vote = clone(golden, root, "decision_vote_order_wrong")
        document = Document(wrong_vote / OUTPUT_NAME)
        document.tables[3].cell(1, 4).text = "4票同意：许承柏、陈予宁、侯景澜、林若棠"
        document.save(wrong_vote / OUTPUT_NAME)
        cases[wrong_vote.name] = score(wrong_vote)

        fabricated = clone(golden, root, "fabricated_assignee_deadline")
        document = Document(fabricated / OUTPUT_NAME)
        document.tables[3].cell(1, 6).text = "魏预"
        document.tables[3].cell(1, 7).text = "2026年3月31日"
        document.save(fabricated / OUTPUT_NAME)
        cases[fabricated.name] = score(fabricated)

        nested = clone(golden, root, "nested_table_reintroduced")
        document = Document(nested / OUTPUT_NAME)
        document.tables[0].cell(0, 0).add_table(rows=1, cols=1).cell(0, 0).text = "错误嵌套"
        document.save(nested / OUTPUT_NAME)
        cases[nested.name] = score(nested)

        flattened = root / "flattened_screenshot_only"
        flattened.mkdir()
        pixel = root / "pixel.png"
        pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document()
        document.add_heading("部门协调会记录整理与核查", 0)
        document.add_picture(str(pixel), width=Inches(6))
        document.save(flattened / OUTPUT_NAME)
        cases[flattened.name] = score(flattened)

        section = clone(golden, root, "landscape_appendix_destroyed")
        document = Document(section / OUTPUT_NAME)
        document.sections[1].orientation = WD_ORIENT.PORTRAIT
        document.sections[1].page_width, document.sections[1].page_height = document.sections[1].page_height, document.sections[1].page_width
        document.save(section / OUTPUT_NAME)
        cases[section.name] = score(section)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite_docx(medium / OUTPUT_NAME, remove_navigation)
        cases[medium.name] = score(medium)

        assert cases["missing_output"] == 0.0, cases
        assert cases["corrupt_docx"] == 0.0, cases
        assert cases["extra_artifact"] == 0.0, cases
        assert 0.0 < cases["protected_location_changed"] <= 0.35, cases
        assert 0.0 < cases["transcript_source_line_missing"] <= 0.55, cases
        assert 0.0 < cases["count_revision_accepted"] <= 0.55, cases
        assert 0.0 < cases["count_changed_to_six"] <= 0.55, cases
        assert 0.0 < cases["count_comment_anchor_removed"] <= 0.55, cases
        assert 0.0 < cases["blank_transcript_row_added"] <= 0.65, cases
        assert 0.0 < cases["first_rule_thick_red"] <= 0.65, cases
        assert 0.0 < cases["deviation_owner_swapped"] <= 0.70, cases
        assert 0.0 < cases["decision_vote_order_wrong"] <= 0.75, cases
        assert 0.0 < cases["fabricated_assignee_deadline"] <= 0.75, cases
        assert 0.0 < cases["nested_table_reintroduced"] <= 0.65, cases
        assert cases["flattened_screenshot_only"] <= 0.35, cases
        assert 0.0 < cases["landscape_appendix_destroyed"] < 1.0, cases
        assert 0.0 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--retain-dir", type=Path)
    arguments = parser.parse_args()
    run(arguments.retain_dir)
