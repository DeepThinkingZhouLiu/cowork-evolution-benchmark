#!/usr/bin/env python3
"""受控购物清单题的任务特有负例校准。"""

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
OUTPUT_NAME = "controlled_shopping_purchase_brief.docx"
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


def table(root, header: str):
    matches = root.xpath("//w:tbl[normalize-space(string(w:tr[1]/w:tc[1])) = $header]", namespaces=NS, header=header)
    assert len(matches) == 1, (header, len(matches))
    return matches[0]


def set_cell(root, header: str, row_number: int, cell_number: int, value: str) -> None:
    target = table(root, header)
    rows = target.xpath("./w:tr", namespaces=NS)
    cells = rows[row_number].xpath("./w:tc", namespaces=NS)
    texts = cells[cell_number].xpath(".//w:t", namespaces=NS)
    assert texts
    texts[0].text = value
    for node in texts[1:]:
        node.text = ""


def delete_row(root, header: str, row_number: int) -> None:
    target = table(root, header)
    rows = target.xpath("./w:tr", namespaces=NS)
    target.remove(rows[row_number])


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
    for field in root.xpath("//w:fldSimple[contains(@w:instr, 'TOC')]", namespaces=NS):
        field.getparent().remove(field)


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="shopping-control-negative-") as temporary:
        root = Path(temporary)
        fresh = root / "fresh"
        fresh.mkdir()
        subprocess.run([PYTHON, str(TASK_ROOT / "solution/solve.py"), "--input-dir", str(TASK_ROOT / "data/input_files"), "--output-dir", str(fresh)], check=True, timeout=180)
        golden = fresh / OUTPUT_NAME
        cases: dict[str, float] = {}

        missing = root / "missing_output"; missing.mkdir(); cases[missing.name] = score(missing)
        corrupt = root / "corrupt_docx"; corrupt.mkdir(); (corrupt / OUTPUT_NAME).write_bytes(b"not a docx"); cases[corrupt.name] = score(corrupt)
        extra = clone(golden, root, "extra_artifact"); (extra / "unchecked-list.txt").write_text("extra", encoding="utf-8"); cases[extra.name] = score(extra)
        wrong_name = root / "wrong_filename"; wrong_name.mkdir(); shutil.copy2(golden, wrong_name / "shopping_list.docx"); cases[wrong_name.name] = score(wrong_name)

        apple = clone(golden, root, "t01_applied_twice")
        rewrite(apple / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: set_cell(xml, "Item", 1, 1, "0.5"))); cases[apple.name] = score(apple)

        onion = clone(golden, root, "removed_onion_kept_current")
        rewrite(onion / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: set_cell(xml, "Item", 3, 0, "onion"))); cases[onion.name] = score(onion)

        cost = clone(golden, root, "wrong_yogurt_cost")
        rewrite(cost / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: set_cell(xml, "Item", 3, 3, "$4.00"))); cases[cost.name] = score(cost)

        ticket = clone(golden, root, "t02_change_row_missing")
        rewrite(ticket / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: delete_row(xml, "Ticket", 2))); cases[ticket.name] = score(ticket)

        order = clone(golden, root, "ticket_order_wrong")
        def reverse_tickets(xml):
            target = table(xml, "Ticket")
            rows = target.xpath("./w:tr", namespaces=NS)
            target.remove(rows[1]); target.insert(1, rows[3])
        rewrite(order / OUTPUT_NAME, lambda parts: document_mutation(parts, reverse_tickets)); cases[order.name] = score(order)

        state = clone(golden, root, "onion_state_unchanged")
        rewrite(state / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: set_cell(xml, "Baseline row", 3, 5, "unchanged"))); cases[state.name] = score(state)

        reconciliation = clone(golden, root, "yogurt_reconciliation_missing")
        rewrite(reconciliation / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: delete_row(xml, "Baseline row", 4))); cases[reconciliation.name] = score(reconciliation)

        calculation = clone(golden, root, "wrong_total_cost")
        rewrite(calculation / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: set_cell(xml, "Control", 5, 2, "$9.00"))); cases[calculation.name] = score(calculation)

        source = clone(golden, root, "pricing_source_missing")
        rewrite(source / OUTPUT_NAME, lambda parts: document_mutation(parts, lambda xml: delete_row(xml, "Frozen source", 5))); cases[source.name] = score(source)

        comments = clone(golden, root, "native_comments_unanchored")
        def strip_comment_anchors(parts):
            document_mutation(parts, lambda xml: [node.getparent().remove(node) for node in xml.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=NS)])
        rewrite(comments / OUTPUT_NAME, strip_comment_anchors); cases[comments.name] = score(comments)

        fields = clone(golden, root, "page_fields_removed")
        def strip_footer_fields(parts):
            for name in [item for item in parts if item.startswith("word/footer") and item.endswith(".xml")]:
                footer = etree.fromstring(parts[name])
                for node in footer.xpath("//w:fldSimple | //w:instrText", namespaces=NS):
                    node.getparent().remove(node)
                parts[name] = etree.tostring(footer, xml_declaration=True, encoding="UTF-8", standalone="yes")
        rewrite(fields / OUTPUT_NAME, strip_footer_fields); cases[fields.name] = score(fields)

        portrait = clone(golden, root, "control_tables_not_landscape")
        def remove_landscape(xml):
            for section in xml.xpath("//w:sectPr", namespaces=NS):
                for size in section.xpath("./w:pgSz", namespaces=NS):
                    size.attrib.pop(f"{{{WNS}}}orient", None)
                    size.set(f"{{{WNS}}}w", "11906")
                    size.set(f"{{{WNS}}}h", "16838")
        rewrite(portrait / OUTPUT_NAME, lambda parts: document_mutation(parts, remove_landscape)); cases[portrait.name] = score(portrait)

        fabricated = clone(golden, root, "fabricated_vendor_and_tax")
        rewrite(fabricated / OUTPUT_NAME, lambda parts: replace_all(parts, "Handoff control: buy only", "Store: QuickMart. Tax rate: 8%. Handoff control: buy only")); cases[fabricated.name] = score(fabricated)

        flattened = root / "flattened_screenshot_only"; flattened.mkdir()
        pixel = root / "pixel.png"
        pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        screenshot = Document(); screenshot.add_heading("Controlled Shopping List", 0); screenshot.add_picture(str(pixel), width=Inches(6)); screenshot.save(flattened / OUTPUT_NAME); cases[flattened.name] = score(flattened)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite(medium / OUTPUT_NAME, lambda parts: document_mutation(parts, remove_navigation)); cases[medium.name] = score(medium)

        assert all(cases[name] == 0 for name in ["missing_output", "corrupt_docx", "extra_artifact", "wrong_filename"]), cases
        directed = [
            "t01_applied_twice", "removed_onion_kept_current", "wrong_yogurt_cost", "t02_change_row_missing",
            "ticket_order_wrong", "onion_state_unchanged", "yogurt_reconciliation_missing", "wrong_total_cost",
            "pricing_source_missing", "native_comments_unanchored", "page_fields_removed", "control_tables_not_landscape",
            "fabricated_vendor_and_tax", "flattened_screenshot_only",
        ]
        for name in directed:
            assert 0 <= cases[name] < .8, (name, cases[name])
        assert .8 < cases["medium_missing_navigation"] < 1.0, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--retain-dir", type=Path)
    arguments = parser.parse_args()
    run(arguments.retain_dir)
