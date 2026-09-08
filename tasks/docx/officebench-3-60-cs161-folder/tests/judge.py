from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from statistics import mean, median

from docx import Document
from docx.enum.section import WD_ORIENT
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-3-60-cs161-folder"
RUBRIC_VERSION = "officebench-3-60-cs161-folder-v0.3"
JUDGE_VERSION = "officebench-3-60-cs161-folder-deterministic-v0.3"
OUTPUT_PATH = Path("CS161/students.docx")
WEIGHTS = {"R001": .10, "R002": .18, "R003": .17, "R004": .18, "R005": .14, "R006": .10, "R007": .08, "R008": .05}
ASSESSMENTS = ["HW1", "HW2", "Midterm", "HW3", "Final"]
BOUNDARY_NOTE = "Verification flags are mechanical source checks only. They do not assign a course grade, academic consequence, misconduct finding, or advising decision."
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W_NS, "cp": CP_NS}


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value if value is not None else "")).strip()


def fraction(correct: int, total: int) -> float:
    return correct / total if total else 0.0


def result_item(criterion_id: str, score: float, evidence: str) -> dict[str, object]:
    return {"criterion_id": criterion_id, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def load_policy() -> dict[str, object]:
    return json.loads(Path(__file__).with_name("review-policy.json").read_text(encoding="utf-8"))


def load_records(path: Path, policy: dict[str, object]) -> tuple[list[dict[str, object]], str, list[str]]:
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    workbook = load_workbook(path, data_only=True, read_only=True)
    sheet_names = workbook.sheetnames
    sheet = workbook[policy["source"]["worksheet"]]
    values = list(sheet.iter_rows(values_only=True))
    workbook.close()
    headers = [clean(value) for value in values[0]]
    records: list[dict[str, object]] = []
    for source_row, row in enumerate(values[1:], 2):
        record = dict(zip(headers, row))
        record["source_row"] = source_row
        scores = [int(record[name]) for name in ASSESSMENTS]
        flags: list[str] = []
        if min(scores) == 0:
            flags.append("ZERO")
        if max(scores) == 100:
            flags.append("PERFECT")
        if max(scores) - min(scores) >= 80:
            flags.append("SPREAD80")
        if abs(int(record["Final"]) - int(record["Midterm"])) >= 60:
            flags.append("SWING60")
        record["scores"] = scores
        record["mean"] = f"{mean(scores):.1f}"
        record["spread"] = max(scores) - min(scores)
        record["delta"] = int(record["Final"]) - int(record["Midterm"])
        record["flags"] = flags
        records.append(record)
    return records, digest, sheet_names


def queue_for(records: list[dict[str, object]]) -> list[dict[str, object]]:
    output = [record for record in records if record["flags"]]
    output.sort(key=lambda record: (-len(record["flags"]), -int(record["spread"]), int(record["source_row"])))
    return output


def block_digest(records: list[dict[str, object]], headers: list[str]) -> str:
    payload = [[record["source_row"], *[record[name] for name in headers]] for record in records]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")).hexdigest()[:12]


def ooxml_audit(path: Path) -> tuple[bool, str, dict[str, bytes]]:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False, "missing or invalid DOCX", {}
    try:
        with zipfile.ZipFile(path) as archive:
            entries = {name: archive.read(name) for name in archive.namelist()}
        prohibited = [name for name in entries if "embeddings/" in name.lower() or name.lower().endswith((".bin", ".exe", ".docm", ".xlsm"))]
        media = [name for name in entries if name.startswith("word/media/")]
        external: list[str] = []
        for name, payload in entries.items():
            if name.endswith(".rels"):
                root = etree.fromstring(payload)
                external.extend(root.xpath("//*[local-name()='Relationship'][@TargetMode='External']/@Target"))
        ok = not prohibited and not media and not external and "word/document.xml" in entries
        return ok, f"parts={len(entries)}; prohibited={prohibited}; media={media}; external={external}", entries
    except Exception as exc:
        return False, f"OOXML parse failed: {type(exc).__name__}: {exc}", {}


def table_matrix(table) -> list[list[str]]:
    return [[clean(cell.text) for cell in row.cells] for row in table.rows]


def find_tables(document: Document, header: list[str]) -> list[list[list[str]]]:
    return [matrix for table in document.tables if (matrix := table_matrix(table)) and matrix[0] == header]


def matrix_score(actual: list[list[str]], expected: list[list[object]]) -> tuple[float, str]:
    target = [[clean(value) for value in row] for row in expected]
    total = sum(len(row) for row in target)
    correct = 0
    for row_index in range(min(len(actual), len(target))):
        for col_index in range(min(len(actual[row_index]), len(target[row_index]))):
            correct += actual[row_index][col_index] == target[row_index][col_index]
    score = fraction(correct, total)
    if len(actual) != len(target):
        score *= fraction(min(len(actual), len(target)), max(len(actual), len(target)))
    return score, f"rows={len(actual)}/{len(target)}; cells={correct}/{total}; exact={actual == target}"


def exact_table(document: Document, header: list[str], expected: list[list[object]]) -> tuple[float, str]:
    matches = find_tables(document, header)
    if len(matches) != 1:
        return 0.0, f"{header}: count={len(matches)}"
    score, evidence = matrix_score(matches[0][1:], expected)
    return score, f"{header}: {evidence}"


def custom_properties(entries: dict[str, bytes]) -> dict[str, str]:
    payload = entries.get("docProps/custom.xml")
    if not payload:
        return {}
    root = etree.fromstring(payload)
    return {node.get("name"): clean("".join(node.xpath(".//text()"))) for node in root.xpath("//cp:property", namespaces=NS)}


def render(path: Path) -> tuple[bool, str]:
    try:
        with tempfile.TemporaryDirectory(prefix="judge-cs161-register-") as temp_dir:
            temp = Path(temp_dir)
            profile = (temp / "lo-profile").resolve()
            subprocess.run(["libreoffice", "--headless", f"-env:UserInstallation={profile.as_uri()}", "--convert-to", "pdf", "--outdir", str(temp), str(path)], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=150)
            pdf = temp / "students.pdf"
            info = subprocess.check_output(["pdfinfo", str(pdf)], text=True, timeout=30)
            pages = int(re.search(r"^Pages:\s+(\d+)", info, re.MULTILINE).group(1))
            expected_titles = [
                "CS161 Roster Verification Register",
                "Assessment Profile & Standing Reconciliation",
                "Canonical Roster • Block B1 • Source rows 2–26",
                "Canonical Roster • Block B2 • Source rows 27–51",
                "Canonical Roster • Block B3 • Source rows 52–76",
                "Canonical Roster • Block B4 • Source rows 77–101",
                "Canonical Roster • Block B5 • Source rows 102–124",
                "Mechanical Verification Queue • Part 1 • Ranks 1–20",
                "Mechanical Verification Queue • Part 2 • Ranks 21–40",
                "Mechanical Verification Queue • Part 3 • Ranks 41–56",
                "Closure & Control Ledger",
            ]
            title_hits = []
            orientations = []
            nonblank = []
            for page_number, expected in enumerate(expected_titles, 1):
                text = clean(subprocess.check_output(["pdftotext", "-f", str(page_number), "-l", str(page_number), "-layout", str(pdf), "-"], text=True, timeout=30))
                title_hits.append(expected in text)
                nonblank.append(bool(text))
                page_info = subprocess.check_output(["pdfinfo", "-f", str(page_number), "-l", str(page_number), str(pdf)], text=True, timeout=30)
                match = re.search(r"(?:Page\s+\d+\s+)?size:\s+([\d.]+)\s+x\s+([\d.]+)\s+pts", page_info)
                if not match:
                    match = re.search(r"Page\s+\d+\s+size:\s+([\d.]+)\s+x\s+([\d.]+)\s+pts", page_info)
                width, height = map(float, match.groups())
                orientations.append("L" if width > height else "P")
            expected_orientations = ["P", *(["L"] * 9), "P"]
            ok = pages == 11 and all(title_hits) and all(nonblank) and orientations == expected_orientations
            return ok, f"pages={pages}; titles={sum(title_hits)}/11; nonblank={sum(nonblank)}/11; orientations={orientations}"
    except Exception as exc:
        return False, f"render failed: {type(exc).__name__}: {exc}"


def evaluate(artifact_dir: Path) -> tuple[list[dict[str, object]], Path]:
    task_root = Path(__file__).parents[1]
    policy = load_policy()
    source = task_root / "data/input_files" / policy["source"]["filename"]
    output = artifact_dir / OUTPUT_PATH
    try:
        records, input_digest, sheet_names = load_records(source, policy)
    except Exception as exc:
        return [result_item(cid, 0, f"source parse failed: {type(exc).__name__}: {exc}" if cid == "R001" else "source unavailable") for cid in WEIGHTS], output
    safe, ooxml_evidence, entries = ooxml_audit(output)
    expected_files = sorted(str(path.relative_to(artifact_dir)) for path in artifact_dir.rglob("*.docx")) if artifact_dir.is_dir() else []
    input_ok = input_digest == policy["source"]["sha256"] and sheet_names == [policy["source"]["worksheet"]] and len(records) == 123
    if not safe:
        return [result_item(cid, 0, f"{ooxml_evidence}; input={input_digest}; files={expected_files}" if cid == "R001" else "valid canonical register unavailable") for cid in WEIGHTS], output
    try:
        document = Document(output)
        native_ok = len(document.tables) == 17 and not document.inline_shapes
        r1 = input_ok and expected_files == [str(OUTPUT_PATH)] and native_ok
        results = [result_item("R001", 1 if r1 else 0, f"{ooxml_evidence}; input={input_digest}; sheet={sheet_names}; rows={len(records)}; files={expected_files}; tables={len(document.tables)}; inline_shapes={len(document.inline_shapes)}")]

        roster_header = ["Row", "Student ID", "Last", "First", "Standing", "HW1", "HW2", "Midterm", "HW3", "Final"]
        roster_tables = find_tables(document, roster_header)
        actual_roster = [row for matrix in roster_tables for row in matrix[1:]]
        expected_roster = [[record["source_row"], record["Student ID"], record["Last Name"], record["First Name"], record["Academic Standing"], *[record[name] for name in ASSESSMENTS]] for record in records]
        identity_actual = [[row[index] for index in (0, 1, 2, 3)] for row in actual_roster]
        identity_expected = [[row[index] for index in (0, 1, 2, 3)] for row in expected_roster]
        identity_score, identity_evidence = matrix_score(identity_actual, identity_expected)
        expected_block_counts = [block["source_row_end"] - block["source_row_start"] + 1 for block in policy["roster_blocks"]]
        actual_block_counts = [len(matrix) - 1 for matrix in roster_tables]
        block_ok = len(roster_tables) == 5 and actual_block_counts == expected_block_counts
        results.append(result_item("R002", (identity_score + int(block_ok)) / 2, f"identity[{identity_evidence}]; tables={len(roster_tables)}; blocks={actual_block_counts}/{expected_block_counts}"))
        source_actual = [[row[index] for index in (4, 5, 6, 7, 8, 9)] for row in actual_roster]
        source_expected = [[row[index] for index in (4, 5, 6, 7, 8, 9)] for row in expected_roster]
        source_score, source_evidence = matrix_score(source_actual, source_expected)
        results.append(result_item("R003", source_score, source_evidence))

        queue = queue_for(records)
        queue_header = ["Rank", "Row", "Student ID", "Name", "Hits", "Mean", "Spread", "Final−Mid", "Rule codes"]
        queue_tables = find_tables(document, queue_header)
        actual_queue = [row for matrix in queue_tables for row in matrix[1:]]
        expected_queue = [[rank, record["source_row"], record["Student ID"], f"{record['First Name']} {record['Last Name']}", len(record["flags"]), record["mean"], record["spread"], f"{int(record['delta']):+d}", " · ".join(record["flags"])] for rank, record in enumerate(queue, 1)]
        queue_score, queue_evidence = matrix_score(actual_queue, expected_queue)
        queue_counts = [len(matrix) - 1 for matrix in queue_tables]
        partition_ok = len(queue_tables) == 3 and queue_counts == policy["queue_parts"] and len(queue) == 56
        results.append(result_item("R004", (queue_score + int(partition_ok)) / 2, f"queue[{queue_evidence}]; tables={len(queue_tables)}; parts={queue_counts}/{policy['queue_parts']}"))

        overall_expected = [[name, f"{mean([int(record[name]) for record in records]):.2f}", f"{median([int(record[name]) for record in records]):.2f}", min(int(record[name]) for record in records), max(int(record[name]) for record in records), sum(int(record[name]) == 0 for record in records), sum(int(record[name]) == 100 for record in records)] for name in ASSESSMENTS]
        standing_expected = []
        for standing in policy["standing_order"]:
            subset = [record for record in records if record["Academic Standing"] == standing]
            standing_expected.append([standing, len(subset), *[f"{mean([int(record[name]) for record in subset]):.2f}" for name in ASSESSMENTS], sum(bool(record["flags"]) for record in subset)])
        codes = [rule["code"] for rule in policy["verification_rules"]]
        incidence_expected = []
        for standing in policy["standing_order"]:
            subset = [record for record in records if record["Academic Standing"] == standing]
            incidence_expected.append([standing, *[sum(code in record["flags"] for record in subset) for code in codes], sum(bool(record["flags"]) for record in subset)])
        incidence_expected.append(["All students", *[sum(code in record["flags"] for record in records) for code in codes], len(queue)])
        aggregate_checks = [
            exact_table(document, ["Assessment", "Mean", "Median", "Minimum", "Maximum", "Zero count", "100 count"], overall_expected),
            exact_table(document, ["Standing", "Count", *[f"{name} mean" for name in ASSESSMENTS], "Flagged"], standing_expected),
            exact_table(document, ["Standing", *codes, "Any flag"], incidence_expected),
        ]
        results.append(result_item("R005", mean(score for score, _ in aggregate_checks), " | ".join(evidence for _, evidence in aggregate_checks)))

        headers = policy["source"]["required_headers"]
        block_ledger = []
        for block in policy["roster_blocks"]:
            subset = [record for record in records if block["source_row_start"] <= record["source_row"] <= block["source_row_end"]]
            block_ledger.append([block["id"], f"{block['source_row_start']}–{block['source_row_end']}", len(subset), subset[0]["Student ID"], subset[-1]["Student ID"], block_digest(subset, headers)])
        queue_ledger = []
        cursor = 0
        for part, count in enumerate(policy["queue_parts"], 1):
            subset = queue[cursor:cursor + count]
            queue_ledger.append([part, f"{cursor + 1}–{cursor + len(subset)}", len(subset), subset[0]["Student ID"], subset[-1]["Student ID"]])
            cursor += count
        reconciliation = [
            ["C01", "Input student rows", 123, 123, "MATCH"], ["C02", "Unique Student IDs", 123, 123, "MATCH"],
            ["C03", "Roster block rows", 123, 123, "MATCH"], ["C04", "Standing count sum", 123, 123, "MATCH"],
            ["C05", "Verification rows", 56, 56, "MATCH"], ["C06", "Queue partition rows", 56, 56, "MATCH"],
            ["C07", "First / last source IDs", "S1000 / S1122", "S1000 / S1122", "MATCH"], ["C08", "Output DOCX count", 1, 1, "MATCH"],
        ]
        expected_props = {
            "SourceSHA256": input_digest, "RosterCount": "123", "VerificationCount": "56", "FirstStudentID": "S1000",
            "LastStudentID": "S1122", "SourceSheet": "Sheet1", "RulesetVersion": policy["schema_version"],
        }
        property_rows = [[key, value] for key, value in expected_props.items()]
        custom = custom_properties(entries)
        property_score = fraction(sum(custom.get(key) == value for key, value in expected_props.items()), len(expected_props))
        control_checks = [
            exact_table(document, ["Block", "Source rows", "Count", "First ID", "Last ID", "SHA-12"], block_ledger),
            exact_table(document, ["Part", "Ranks", "Count", "First ID", "Last ID"], queue_ledger),
            exact_table(document, ["Check", "Control", "Expected", "Observed", "Result"], reconciliation),
            exact_table(document, ["Property", "Value"], property_rows),
            (property_score, f"custom properties={sum(custom.get(key) == value for key, value in expected_props.items())}/{len(expected_props)}; actual={custom}"),
        ]
        results.append(result_item("R006", mean(score for score, _ in control_checks), " | ".join(evidence for _, evidence in control_checks)))

        document_root = etree.fromstring(entries["word/document.xml"])
        bookmarks = document_root.xpath("//w:bookmarkStart/@w:name", namespaces=NS)
        links = document_root.xpath("//w:hyperlink/@w:anchor", namespaces=NS)
        expected_bookmarks = ["control_page", "assessment_profile", "roster_b1", "roster_b2", "roster_b3", "roster_b4", "roster_b5", "queue_1", "queue_2", "queue_3", "closure"]
        expected_links = expected_bookmarks[1:]
        header_rows = len(document_root.xpath("//w:tr[w:trPr/w:tblHeader]", namespaces=NS))
        no_split = len(document_root.xpath("//w:tr[w:trPr/w:cantSplit]", namespaces=NS))
        total_rows = len(document_root.xpath("//w:tr", namespaces=NS))
        footer_fields = []
        for name, payload in entries.items():
            if name.startswith("word/footer") and name.endswith(".xml"):
                root = etree.fromstring(payload)
                footer_fields.extend(clean(value) for value in root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
        heading_texts = [clean(paragraph.text) for paragraph in document.paragraphs if paragraph.style.name.startswith("Heading")]
        page_headings = ["Assessment Profile & Standing Reconciliation", *[f"Canonical Roster • Block B{index} • Source rows {policy['roster_blocks'][index - 1]['source_row_start']}–{policy['roster_blocks'][index - 1]['source_row_end']}" for index in range(1, 6)], *[f"Mechanical Verification Queue • Part {index} • Ranks {start}–{end}" for index, (start, end) in enumerate(((1, 20), (21, 40), (41, 56)), 1)], "Closure & Control Ledger"]
        orientations = [section.orientation for section in document.sections]
        header_text = [clean(section.header.paragraphs[0].text) for section in document.sections]
        structure_checks = [
            len(document.sections) == 3 and orientations == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT],
            len(document.tables) == 17,
            bookmarks == expected_bookmarks,
            links == expected_links,
            header_rows == 17,
            no_split >= total_rows - 1,
            footer_fields.count("PAGE") == 3 and footer_fields.count("NUMPAGES") == 3,
            all(title in heading_texts for title in page_headings),
            header_text == ["CS161 • CONTROL & NAVIGATION", "CS161 • SOURCE REGISTER & VERIFICATION QUEUE", "CS161 • CLOSURE & CONTROL LEDGER"],
        ]
        results.append(result_item("R007", fraction(sum(structure_checks), len(structure_checks)), f"checks={sum(structure_checks)}/{len(structure_checks)}; sections={orientations}; tables={len(document.tables)}; bookmarks={bookmarks}; links={links}; header_rows={header_rows}; no_split={no_split}/{total_rows}; fields={footer_fields}; headers={header_text}"))

        render_ok, render_evidence = render(output)
        text = clean(" ".join(paragraph.text for paragraph in document.paragraphs) + " " + " ".join(cell.text for table in document.tables for row in table.rows for cell in row.cells))
        unsupported_phrases = [phrase for phrase in ("assigned final grade", "failed the course", "disciplinary action", "misconduct confirmed", "academic probation", "must retake") if phrase in text.casefold()]
        boundary_count = text.count(BOUNDARY_NOTE)
        results.append(result_item("R008", (int(render_ok) + int(boundary_count >= 2 and not unsupported_phrases)) / 2, f"{render_evidence}; boundary_count={boundary_count}; unsupported={unsupported_phrases}"))
        return results, output
    except Exception as exc:
        return [result_item(cid, 0, f"parse error: {type(exc).__name__}: {exc}" if cid == "R001" else "parse error; not scored") for cid in WEIGHTS], output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path, required=True)
    parser.add_argument("--reward-file", type=Path, required=True)
    args = parser.parse_args()
    results, output = evaluate(args.artifact_dir)
    scores = {entry["criterion_id"]: float(entry["score"]) for entry in results}
    reward = sum(WEIGHTS[criterion_id] * scores.get(criterion_id, 0) for criterion_id in WEIGHTS)
    for criterion_id, cap in (("R001", 0.0), ("R002", .30), ("R003", .46), ("R004", .60), ("R005", .72), ("R006", .78), ("R007", .83)):
        if scores.get(criterion_id, 0) < 1:
            reward = min(reward, cap)
    reward = round(reward, 6)
    passed = reward >= .86 and all(scores.get(criterion_id, 0) == 1 for criterion_id in WEIGHTS)
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": passed,
    }
    args.result.parent.mkdir(parents=True, exist_ok=True)
    args.reward_file.parent.mkdir(parents=True, exist_ok=True)
    args.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    args.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if passed else 1)


if __name__ == "__main__":
    main()
