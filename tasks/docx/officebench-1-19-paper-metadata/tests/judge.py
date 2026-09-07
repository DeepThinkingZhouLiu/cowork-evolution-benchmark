from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree


WEIGHTS = {"R001": 8, "R002": 14, "R003": 15, "R004": 18, "R005": 10, "R006": 13, "R007": 8, "R008": 5, "R009": 4, "R010": 4, "R011": 1}
HURDLES = {"R001": 0, "R002": 0.35, "R003": 0.55, "R004": 0.65, "R005": 0.72, "R006": 0.72, "R007": 0.78}
W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def criterion(identifier: str, score: float, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": round(score, 6), "evidence": evidence, "method": "deterministic", "raw": {"weight": WEIGHTS[identifier]}}


def normalized(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def source_truth(root: Path) -> tuple[str, str]:
    pdf = root / "paper.pdf"
    text = " ".join(subprocess.check_output(["pdftotext", "-layout", str(pdf), "-"], text=True).split())
    info = subprocess.check_output(["pdfinfo", str(pdf)], text=True)
    required = [
        "OS-Copilot", "Zhiyong Wu", "Chengcheng Han", "Zichen Ding", "Zhenmin Weng",
        "Zhoumianze Liu", "Shunyu Yao", "Tao Yu", "Lingpeng Kong", "Shanghai AI Laboratory",
        "East China Normal University", "Princeton University", "The University of Hong Kong",
        "wuzhiyong@pjlab.org.cn", "2402.07456v2", "15 Feb 2024", "outperforms previous methods by 35%",
        "focused working environment", "drawing a chart in Excel", "creating a website for OS-Copilot",
    ]
    assert all(value in text for value in required)
    assert "Pages:           1" in info and "Tagged:          no" in info
    return text, info


def evaluate(source_dir: Path, artifact_dir: Path) -> list[dict]:
    files = sorted(path.name for path in artifact_dir.iterdir()) if artifact_dir.is_dir() else []
    output = artifact_dir / "os_copilot_research_intake.docx"
    valid = files == [output.name] and output.is_file() and zipfile.is_zipfile(output)
    results = [criterion("R001", 1 if valid else 0, f"files={files}; native_docx={valid}")]
    if not valid:
        return results
    source_truth(source_dir)
    document = Document(output)
    text = normalized("\n".join(paragraph.text for paragraph in document.paragraphs) + "\n" + "\n".join(cell.text for table in document.tables for row in table.rows for cell in row.cells))

    identity = [
        "OS-Copilot: Towards Generalist Computer Agents with Self-Improvement", "arXiv:2402.07456v2 [cs.AI]",
        "15 Feb 2024", "wuzhiyong@pjlab.org.cn", "https://os-copilot.github.io/", "1 page; Letter; PDF 1.5",
        "Both blank in PDF properties", "Tagged: no",
    ]
    identity_count = sum(value in text for value in identity)
    results.append(criterion("R002", identity_count / len(identity), f"identity/properties={identity_count}/{len(identity)}"))

    authors = ["Zhiyong Wu", "Chengcheng Han", "Zichen Ding", "Zhenmin Weng", "Zhoumianze Liu", "Shunyu Yao", "Tao Yu", "Lingpeng Kong"]
    author_table = next((table for table in document.tables if len(table.rows) == 9 and len(table.columns) == 4), None)
    displayed = [] if author_table is None else [row.cells[1].text.strip() for row in author_table.rows[1:]]
    affiliations = [
        "Zhiyong Wu|Shanghai AI Laboratory|Equal contribution; listed contact",
        "Chengcheng Han|East China Normal University|Equal contribution",
        "Zichen Ding|East China Normal University|—", "Zhenmin Weng|East China Normal University|—",
        "Zhoumianze Liu|Shanghai AI Laboratory|—", "Shunyu Yao|Princeton University|—",
        "Tao Yu|The University of Hong Kong|—", "Lingpeng Kong|The University of Hong Kong|—",
    ]
    actual = [] if author_table is None else ["|".join(cell.text.strip() for cell in row.cells[1:4]) for row in author_table.rows[1:]]
    boundaries = all(value in text for value in ["Only Zhiyong Wu and Chengcheng Han carry the equal-contribution marker.", "does not label a corresponding author"])
    authorship_score = (int(displayed == authors) + sum(value in actual for value in affiliations) / 8 + int(boundaries)) / 3
    results.append(criterion("R003", authorship_score, f"order={displayed == authors}; mappings={sum(value in actual for value in affiliations)}/8; boundaries={boundaries}"))

    abstract_facts = [
        "Autonomous computer interaction is longstanding", "Most agents target one software product or website",
        "OS-Copilot supplies the generalist framework", "FRIDAY is its self-improving embodied agent",
        "Web, code terminals, files, multimedia, and third-party applications", "beat previous GAIA methods by 35%",
        "learn Excel and PowerPoint control with minimal supervision", "more capable general-purpose computer agents",
    ]
    abstract_count = sum(value in text for value in abstract_facts)
    results.append(criterion("R004", abstract_count / len(abstract_facts), f"structured abstract facts={abstract_count}/{len(abstract_facts)}"))

    figure_facts = ["Prepare a focused working environment", "Calculate and draw a chart in Excel", "Create a website for OS-Copilot", "Top = initial OS state; bottom = final state", "Bottom text gives FRIDAY's subtasks", "Boxes/ovals highlight changes"]
    figure_count = sum(value in text for value in figure_facts)
    results.append(criterion("R005", figure_count / len(figure_facts), f"figure facts={figure_count}/{len(figure_facts)}"))

    matrix = next((table for table in document.tables if len(table.rows) == 13 and len(table.columns) == 5), None)
    ids = [] if matrix is None else [row.cells[0].text.strip() for row in matrix.rows[1:]]
    locators = [] if matrix is None else [row.cells[3].text.strip() for row in matrix.rows[1:]]
    statuses = [] if matrix is None else [row.cells[4].text.strip() for row in matrix.rows[1:]]
    matrix_score = (int(ids == [f"E{i:02}" for i in range(1, 13)]) + int(all(value.startswith("Page 1") for value in locators)) + int(all(value.startswith("Verified") for value in statuses))) / 3
    results.append(criterion("R006", matrix_score, f"matrix={matrix is not None}; ids={len(ids)}; locators={len(locators)}; statuses={len(statuses)}"))

    boundary = all(value in text for value in ["The frozen PDF contains one page", "Request the complete paper before technical review", "Do not add a DOI, venue, publisher, corresponding-author designation, benchmark baseline, confidence interval"])
    forbidden = ["doi.org/", "NeurIPS", "ICLR", "AAAI", "corresponding author is Zhiyong", "GPT-4 baseline", "95% confidence"]
    fabricated = [value for value in forbidden if value.lower() in text.lower()]
    results.append(criterion("R007", 1 if boundary and not fabricated else 0, f"boundary={boundary}; fabricated={fabricated}"))

    with zipfile.ZipFile(output) as archive:
        xml = etree.fromstring(archive.read("word/document.xml"))
        names = set(xml.xpath("//w:bookmarkStart/@w:name", namespaces=W_NS))
        anchors = xml.xpath("//w:hyperlink/@w:anchor", namespaces=W_NS)
        navigation = sum(f"evidence_E{i:02}" in names and f"evidence_E{i:02}" in anchors for i in range(1, 13))
        results.append(criterion("R008", navigation / 12, f"bookmark-links={navigation}/12"))
        starts = xml.xpath("//w:commentRangeStart/@w:id", namespaces=W_NS)
        ends = xml.xpath("//w:commentRangeEnd/@w:id", namespaces=W_NS)
        references = xml.xpath("//w:commentReference/@w:id", namespaces=W_NS)
        comment_text = ""
        if "word/comments.xml" in archive.namelist():
            comments = etree.fromstring(archive.read("word/comments.xml"))
            comment_text = normalized(" ".join(comments.xpath("//w:t/text()", namespaces=W_NS)))
        comment_markers = ["asterisk appears beside these two names only", "percentage but not a baseline identity", "evidence boundary, not a conclusion"]
        comment_score = (int(starts == ends == references == ["0", "1", "2"]) + sum(value in comment_text for value in comment_markers) / 3) / 2
        results.append(criterion("R009", comment_score, f"comment ids={starts}/{ends}/{references}; notes={sum(value in comment_text for value in comment_markers)}/3"))
        section_count = len(xml.xpath("//w:sectPr", namespaces=W_NS))
        landscape = bool(xml.xpath("//w:sectPr/w:pgSz[@w:orient='landscape']", namespaces=W_NS))
        no_splits = len(xml.xpath("//w:cantSplit", namespaces=W_NS))
        fields = " ".join(" ".join(etree.fromstring(archive.read(name)).xpath("//w:fldSimple/@w:instr | //w:instrText/text()", namespaces=W_NS)) for name in archive.namelist() if name.startswith("word/") and name.endswith(".xml"))
        core = normalized(etree.fromstring(archive.read("docProps/core.xml")).xpath("string(.)"))
    structure = sum([len(document.tables) == 6, section_count >= 3, landscape, "TOC" in fields, "PAGE" in fields and "NUMPAGES" in fields, no_splits >= 35, "OS-Copilot Research Intake & Evidence Review" in core])
    results.append(criterion("R010", structure / 7, f"tables={len(document.tables)}; sections={section_count}; landscape={landscape}; fields={fields}; cantSplit={no_splits}; core={bool(core)}"))
    headings = sum(paragraph.style.name.startswith("Heading") for paragraph in document.paragraphs)
    results.append(criterion("R011", 1 if headings >= 8 and len(text) > 3500 else 0.5, f"headings={headings}; chars={len(text)}"))
    return results


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    arguments = parser.parse_args()
    source_dir = Path(__file__).parents[1] / "data/input_files"
    try:
        results = evaluate(source_dir, arguments.artifact_dir)
    except Exception as error:
        results = [criterion("R001", 0, f"exception={type(error).__name__}: {error}")]
    scores = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[key] * scores.get(key, 0) for key in WEIGHTS) / 100
    failures = [key for key in HURDLES if scores.get(key, 0) < 0.999999]
    if failures:
        reward = min(reward, min(HURDLES[key] for key in failures))
    output = arguments.artifact_dir / "os_copilot_research_intake.docx"
    payload = {
        "task_id": "cowork-evo/officebench-1-19-paper-metadata",
        "rubric_version": "officebench-1-19-paper-metadata-v0.3",
        "judge_version": "officebench-1-19-paper-metadata-deterministic-v0.3",
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": round(reward, 6),
        "passed": reward >= 0.8 and not failures,
        "hurdle_failures": failures,
    }
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    arguments.reward_file.parent.mkdir(parents=True, exist_ok=True)
    arguments.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")


if __name__ == "__main__":
    main()
