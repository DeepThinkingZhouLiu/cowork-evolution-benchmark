#!/usr/bin/env python3
"""并列高分发布审计题的任务特有负例校准。"""

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
from docx.shared import Inches
from lxml import etree


TASK_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_NAME = "highest_score_recognition_review.docx"
PYTHON = sys.executable
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WNS}


def score(case: Path) -> float:
    result = case.parent / f"{case.name}-result.json"
    reward = case.parent / f"{case.name}-reward.txt"
    subprocess.run([PYTHON, str(TASK_ROOT / "tests/judge.py"), "--artifact-dir", str(case), "--result", str(result), "--reward-file", str(reward)], check=False, timeout=360)
    return float(json.loads(result.read_text(encoding="utf-8"))["reward"])


def clone(golden: Path, root: Path, name: str) -> Path:
    case = root / name
    case.mkdir()
    shutil.copy2(golden, case / OUTPUT_NAME)
    return case


def rewrite(path: Path, mutator) -> None:
    temporary = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        infos = {info.filename: info for info in source.infolist()}
    mutator(parts)
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for name, payload in parts.items():
            target.writestr(infos[name] if name in infos else name, payload)
    temporary.replace(path)


def document_mutation(parts: dict[str, bytes], mutator) -> None:
    root = etree.fromstring(parts["word/document.xml"])
    mutator(root)
    parts["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def replace_all(parts: dict[str, bytes], before: str, after: str) -> None:
    count = 0
    for name, payload in list(parts.items()):
        if not name.endswith(".xml"):
            continue
        root = etree.fromstring(payload)
        for node in root.xpath("//w:t", namespaces=NS):
            if before in (node.text or ""):
                node.text = (node.text or "").replace(before, after)
                count += 1
        parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")
    assert count, before


def remove_navigation(root) -> None:
    for node in root.xpath("//w:bookmarkStart | //w:bookmarkEnd", namespaces=NS):
        node.getparent().remove(node)
    for hyperlink in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS):
        parent = hyperlink.getparent()
        position = parent.index(hyperlink)
        for child in list(hyperlink):
            parent.insert(position, child)
            position += 1
        parent.remove(hyperlink)


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="recognition-review-negative-") as temporary:
        root = Path(temporary)
        fresh = root / "fresh"
        fresh.mkdir()
        subprocess.run([PYTHON, str(TASK_ROOT / "solution/solve.py"), "--input-dir", str(TASK_ROOT / "data/input_files"), "--output-dir", str(fresh)], check=True, timeout=180)
        golden = fresh / OUTPUT_NAME
        cases: dict[str, float] = {}

        missing = root / "missing_output"; missing.mkdir(); cases[missing.name] = score(missing)
        corrupt = root / "corrupt_docx"; corrupt.mkdir(); (corrupt / OUTPUT_NAME).write_bytes(b"not a docx"); cases[corrupt.name] = score(corrupt)
        extra = clone(golden, root, "extra_artifact"); (extra / "release-note.txt").write_text("extra", encoding="utf-8"); cases[extra.name] = score(extra)
        wrong_name = root / "wrong_filename"; wrong_name.mkdir(); shutil.copy2(golden, wrong_name / "Rachel.docx"); cases[wrong_name.name] = score(wrong_name)

        released = clone(golden, root, "false_release_ready")
        rewrite(released / OUTPUT_NAME, lambda parts: replace_all(parts, "HOLD", "READY")); cases[released.name] = score(released)

        omitted = clone(golden, root, "cohighest_tina_omitted")
        rewrite(omitted / OUTPUT_NAME, lambda parts: replace_all(parts, "Tina", "Tony")); cases[omitted.name] = score(omitted)

        sole = clone(golden, root, "rachel_claimed_sole_winner")
        rewrite(sole / OUTPUT_NAME, lambda parts: replace_all(parts, "The nomination shortlist names Rachel at 100 as its highest entry.", "Rachel is the sole winner and is the only highest-scoring student.")); cases[sole.name] = score(sole)

        draft = clone(golden, root, "rachel_draft_wrong_tie_status")
        rewrite(draft / OUTPUT_NAME, lambda parts: replace_all(parts, "Dear Rachel, your Midterm 1 assessment record is 100, tied for the highest", "Dear Rachel, your Midterm 1 assessment record is 100 and is the only highest")); cases[draft.name] = score(draft)

        metric = clone(golden, root, "wrong_midterm1_spread")
        rewrite(metric / OUTPUT_NAME, lambda parts: replace_all(parts, "100 / 60 / 40", "100 / 60 / 35")); cases[metric.name] = score(metric)

        shortlist = clone(golden, root, "shortlist_row_missing")
        def delete_shortlist_row(xml_root):
            tables = xml_root.xpath("//w:tbl[w:tr[1]//w:t[contains(., 'Shortlist row')]]", namespaces=NS)
            rows = tables[0].xpath("./w:tr", namespaces=NS)
            tables[0].remove(rows[-1])
        rewrite(shortlist / OUTPUT_NAME, lambda parts: document_mutation(parts, delete_shortlist_row)); cases[shortlist.name] = score(shortlist)

        ledger = clone(golden, root, "duplicate_alice_row_collapsed")
        def delete_alice_row(xml_root):
            tables = xml_root.xpath("//w:tbl[w:tr[1]//w:t[contains(., 'Source row')]]", namespaces=NS)
            rows = tables[0].xpath("./w:tr", namespaces=NS)
            target = next(row for row in rows[1:] if "Alice" in "".join(row.itertext()))
            tables[0].remove(target)
        rewrite(ledger / OUTPUT_NAME, lambda parts: document_mutation(parts, delete_alice_row)); cases[ledger.name] = score(ledger)

        delta = clone(golden, root, "tina_delta_wrong")
        rewrite(delta / OUTPUT_NAME, lambda parts: replace_all(parts, "-33", "+33")); cases[delta.name] = score(delta)

        comments = clone(golden, root, "draft_comments_removed")
        def strip_comments(parts):
            document_mutation(parts, lambda xml_root: [node.getparent().remove(node) for node in xml_root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=NS)])
        rewrite(comments / OUTPUT_NAME, strip_comments); cases[comments.name] = score(comments)

        fields = clone(golden, root, "page_fields_removed")
        def strip_fields(parts):
            for name in [item for item in parts if item.startswith("word/footer") and item.endswith(".xml")]:
                xml_root = etree.fromstring(parts[name])
                for node in xml_root.xpath("//w:fldSimple | //w:instrText", namespaces=NS):
                    node.getparent().remove(node)
                parts[name] = etree.tostring(xml_root, xml_declaration=True, encoding="UTF-8", standalone="yes")
        rewrite(fields / OUTPUT_NAME, strip_fields); cases[fields.name] = score(fields)

        portrait = clone(golden, root, "ledger_section_portrait")
        def make_portrait(xml_root):
            section = xml_root.xpath("//w:sectPr", namespaces=NS)[-1]
            size = section.xpath("./w:pgSz", namespaces=NS)[0]
            size.attrib.pop(f"{{{WNS}}}orient", None)
            size.set(f"{{{WNS}}}w", "11906")
            size.set(f"{{{WNS}}}h", "16838")
        rewrite(portrait / OUTPUT_NAME, lambda parts: document_mutation(parts, make_portrait)); cases[portrait.name] = score(portrait)

        fabricated = clone(golden, root, "fabricated_identity_and_gpa")
        rewrite(fabricated / OUTPUT_NAME, lambda parts: replace_all(parts, "Resolve Tina's missing nomination", "Student ID: 123456; GPA: 4.0; Tina is eligible for the award. Resolve Tina's missing nomination")); cases[fabricated.name] = score(fabricated)

        flattened = root / "flattened_screenshot_only"; flattened.mkdir()
        pixel = root / "pixel.png"
        pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        screenshot = Document(); screenshot.add_heading("Midterm Recognition Review", 0); screenshot.add_picture(str(pixel), width=Inches(6)); screenshot.save(flattened / OUTPUT_NAME); cases[flattened.name] = score(flattened)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite(medium / OUTPUT_NAME, lambda parts: document_mutation(parts, remove_navigation)); cases[medium.name] = score(medium)

        assert all(cases[name] == 0 for name in ["missing_output", "corrupt_docx", "extra_artifact", "wrong_filename"]), cases
        directed = ["false_release_ready", "cohighest_tina_omitted", "rachel_claimed_sole_winner", "rachel_draft_wrong_tie_status", "wrong_midterm1_spread", "shortlist_row_missing", "duplicate_alice_row_collapsed", "tina_delta_wrong", "draft_comments_removed", "page_fields_removed", "ledger_section_portrait", "fabricated_identity_and_gpa", "flattened_screenshot_only"]
        for name in directed:
            assert 0 <= cases[name] < .8, (name, cases[name])
        assert .8 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--retain-dir", type=Path)
    args = parser.parse_args()
    run(args.retain_dir)
