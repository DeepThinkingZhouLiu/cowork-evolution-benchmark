#!/usr/bin/env python3
"""OS-Copilot 研究入库题的任务特有负例校准。"""

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
OUTPUT_NAME = "os_copilot_research_intake.docx"
PYTHON = sys.executable
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


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


def rewrite(path: Path, mutator) -> None:
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


def remove_comments(root) -> None:
    for node in root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=W_NS):
        node.getparent().remove(node)


def remove_matrix_row(root) -> None:
    matrix = root.xpath("//w:body/w:tbl[3]", namespaces=W_NS)[0]
    rows = matrix.xpath("./w:tr", namespaces=W_NS)
    matrix.remove(rows[7])


def run(retain_dir: Path | None = None) -> None:
    with tempfile.TemporaryDirectory(prefix="paper-metadata-negative-") as temporary:
        root = Path(temporary)
        fresh = root / "fresh"
        fresh.mkdir()
        subprocess.run([PYTHON, str(TASK_ROOT / "solution/solve.py"), "--input-dir", str(TASK_ROOT / "data/input_files"), "--output-dir", str(fresh)], check=True, timeout=180)
        golden = fresh / OUTPUT_NAME
        cases: dict[str, float] = {}

        missing = root / "missing_output"; missing.mkdir(); cases[missing.name] = score(missing)
        corrupt = root / "corrupt_docx"; corrupt.mkdir(); (corrupt / OUTPUT_NAME).write_bytes(b"not a docx"); cases[corrupt.name] = score(corrupt)
        extra = clone(golden, root, "extra_artifact"); (extra / "paper.txt").write_text("extra", encoding="utf-8"); cases[extra.name] = score(extra)

        title = clone(golden, root, "title_truncated")
        replace_all(title / OUTPUT_NAME, "OS-Copilot: Towards Generalist Computer Agents with Self-Improvement", "OS-Copilot")
        cases[title.name] = score(title)

        order = clone(golden, root, "author_order_swapped")
        document = Document(order / OUTPUT_NAME)
        author_table = next(table for table in document.tables if len(table.rows) == 9 and len(table.columns) == 4)
        first, second = author_table.cell(1, 1).text, author_table.cell(2, 1).text
        author_table.cell(1, 1).text, author_table.cell(2, 1).text = second, first
        document.save(order / OUTPUT_NAME)
        cases[order.name] = score(order)

        affiliation = clone(golden, root, "wu_affiliation_wrong")
        document = Document(affiliation / OUTPUT_NAME)
        author_table = next(table for table in document.tables if len(table.rows) == 9 and len(table.columns) == 4)
        author_table.cell(1, 2).text = "Princeton University"
        document.save(affiliation / OUTPUT_NAME)
        cases[affiliation.name] = score(affiliation)

        metric = clone(golden, root, "gaia_metric_changed")
        replace_all(metric / OUTPUT_NAME, "35%", "53%")
        cases[metric.name] = score(metric)

        figure = clone(golden, root, "figure_excel_scenario_missing")
        replace_all(figure / OUTPUT_NAME, "Calculate and draw a chart in Excel", "Prepare a spreadsheet")
        cases[figure.name] = score(figure)

        matrix = clone(golden, root, "evidence_matrix_row_missing")
        rewrite(matrix / OUTPUT_NAME, remove_matrix_row)
        cases[matrix.name] = score(matrix)

        doi = clone(golden, root, "fabricated_doi")
        document = Document(doi / OUTPUT_NAME); document.add_paragraph("DOI: https://doi.org/10.0000/os-copilot"); document.save(doi / OUTPUT_NAME)
        cases[doi.name] = score(doi)

        corresponding = clone(golden, root, "fabricated_corresponding_author")
        document = Document(corresponding / OUTPUT_NAME); document.add_paragraph("Corresponding author is Zhiyong Wu."); document.save(corresponding / OUTPUT_NAME)
        cases[corresponding.name] = score(corresponding)

        count = clone(golden, root, "pdf_page_count_wrong")
        replace_all(count / OUTPUT_NAME, "1 page; Letter; PDF 1.5", "12 pages; Letter; PDF 1.5")
        cases[count.name] = score(count)

        comments = clone(golden, root, "review_comments_removed")
        rewrite(comments / OUTPUT_NAME, remove_comments)
        cases[comments.name] = score(comments)

        flattened = root / "flattened_screenshot_only"; flattened.mkdir()
        pixel = root / "pixel.png"; pixel.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="))
        document = Document(); document.add_heading("OS-Copilot Research Intake & Evidence Review", 0); document.add_picture(str(pixel), width=Inches(6)); document.save(flattened / OUTPUT_NAME)
        cases[flattened.name] = score(flattened)

        landscape = clone(golden, root, "landscape_matrix_destroyed")
        document = Document(landscape / OUTPUT_NAME); document.sections[1].orientation = WD_ORIENT.PORTRAIT; document.sections[1].page_width, document.sections[1].page_height = document.sections[1].page_height, document.sections[1].page_width; document.save(landscape / OUTPUT_NAME)
        cases[landscape.name] = score(landscape)

        medium = clone(golden, root, "medium_missing_navigation")
        rewrite(medium / OUTPUT_NAME, remove_navigation)
        cases[medium.name] = score(medium)

        assert cases["missing_output"] == 0 and cases["corrupt_docx"] == 0 and cases["extra_artifact"] == 0, cases
        assert 0 < cases["title_truncated"] <= 0.35, cases
        assert 0 < cases["author_order_swapped"] <= 0.55, cases
        assert 0 < cases["wu_affiliation_wrong"] <= 0.55, cases
        assert 0 < cases["gaia_metric_changed"] <= 0.65, cases
        assert 0 < cases["figure_excel_scenario_missing"] <= 0.72, cases
        assert 0 < cases["evidence_matrix_row_missing"] <= 0.72, cases
        assert 0 < cases["fabricated_doi"] <= 0.78 and 0 < cases["fabricated_corresponding_author"] <= 0.78, cases
        assert 0 < cases["pdf_page_count_wrong"] <= 0.35, cases
        assert 0 < cases["review_comments_removed"] < 1, cases
        assert cases["flattened_screenshot_only"] <= 0.35, cases
        assert 0 < cases["landscape_matrix_destroyed"] < 1, cases
        assert 0 < cases["medium_missing_navigation"] < 1, cases
        if retain_dir is not None:
            shutil.copytree(root, retain_dir)
        print(json.dumps(cases, ensure_ascii=False, indent=2, sort_keys=True))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--retain-dir", type=Path); arguments = parser.parse_args(); run(arguments.retain_dir)
