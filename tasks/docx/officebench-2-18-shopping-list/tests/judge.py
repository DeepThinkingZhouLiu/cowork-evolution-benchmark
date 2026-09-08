from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from docx import Document
from lxml import etree
from openpyxl import load_workbook


TASK_ID = "cowork-evo/officebench-2-18-shopping-list"
RUBRIC_VERSION = "officebench-2-18-shopping-list-v0.3"
JUDGE_VERSION = "officebench-2-18-shopping-list-deterministic-v0.3"
OUTPUT_NAME = "controlled_shopping_purchase_brief.docx"
WEIGHTS = {"R001": .08, "R002": .18, "R003": .13, "R004": .12, "R005": .11, "R006": .09, "R007": .08, "R008": .06, "R009": .08, "R010": .07}
WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RELNS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": WNS}


@dataclass(frozen=True)
class Ticket:
    identifier: str
    filename: str
    item: str
    action: str
    before: Decimal
    after: Decimal
    source_location: str

    @property
    def delta(self) -> Decimal:
        return self.after - self.before


def criterion(identifier: str, score: float, evidence: str) -> dict:
    return {"criterion_id": identifier, "score": round(max(0.0, min(1.0, score)), 6), "evidence": evidence, "method": "deterministic"}


def number(value: Decimal) -> str:
    return str(int(value)) if value == value.to_integral() else str(value.normalize())


def signed(value: Decimal) -> str:
    return f"+{number(value)}" if value >= 0 else number(value)


def money(value: Decimal) -> str:
    return f"${value:.2f}"


def load_record(path: Path) -> dict:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value.get("task"), str) or not isinstance(value.get("evaluation"), list):
        raise ValueError(f"JSON 记录结构异常: {path.name}")
    return value


def derive_truth(source_dir: Path):
    workbook = load_workbook(source_dir / "baseline_shopping_list.xlsx", data_only=True, read_only=True)
    sheet = workbook.active
    if [cell.value for cell in sheet[1]][:2] != ["Name", "amount"]:
        raise ValueError("基线表头不符合冻结证据")
    baseline = [(index, str(item), Decimal(str(amount))) for index, (item, amount) in enumerate(sheet.iter_rows(min_row=2, max_col=2, values_only=True), start=2)]
    workbook.close()
    if len(baseline) != 3:
        raise ValueError(f"基线行数异常: {len(baseline)}")
    baseline_map = OrderedDict((item, amount) for _, item, amount in baseline)

    t01 = load_record(source_dir / "change_T01_reduce_apple.json")
    if "reduce half amount of apple" not in t01["task"].lower():
        raise ValueError("T01 语义异常")
    apple_target = Decimal(str(t01["evaluation"][0]["args"]["matches"][0]["value"]))
    if apple_target != baseline_map["apple"] / 2:
        raise ValueError("T01 evaluator 与基线减半不一致")

    t02 = load_record(source_dir / "change_T02_remove_onion.json")
    if "remove this entry" not in t02["task"].lower() or "onion" not in t02["task"].lower():
        raise ValueError("T02 语义异常")
    keywords = [str(value) for value in t02["evaluation"][0]["args"]["keywords"]]
    if keywords != ["onion", number(baseline_map["onion"])]:
        raise ValueError("T02 evaluator 与基线不一致")

    t03 = load_record(source_dir / "change_T03_add_yogurt.json")
    matches = t03["evaluation"][0]["args"]["matches"]
    new_item = str(next(match["value"] for match in matches if match["col"] == "1"))
    new_amount = Decimal(str(next(match["value"] for match in matches if match["col"] == "2")))
    if "add to shopping list" not in t03["task"].lower():
        raise ValueError("T03 语义异常")

    pricing = load_record(source_dir / "pricing_policy.json")
    rate_match = re.search(r"apple is\s*([0-9.]+)\$/pound,\s*others are\s*([0-9.]+)\$/pound", pricing["task"], re.I)
    if not rate_match:
        raise ValueError("定价规则不可解析")
    apple_rate, other_rate = map(Decimal, rate_match.groups())

    tickets = [
        Ticket("T01", "change_T01_reduce_apple.json", "apple", "Reduce by half", baseline_map["apple"], apple_target, "evaluation[0].args.matches[0]"),
        Ticket("T02", "change_T02_remove_onion.json", "onion", "Remove — already at home", baseline_map["onion"], Decimal("0"), "task + evaluation[0].args.keywords"),
        Ticket("T03", "change_T03_add_yogurt.json", new_item, "Add approved item", Decimal("0"), new_amount, "evaluation[0].args.matches[0:2]"),
    ]
    current = OrderedDict(baseline_map)
    current["apple"] = apple_target
    del current["onion"]
    current[new_item] = new_amount
    rates = {item: apple_rate if item == "apple" else other_rate for item in current}
    costs = {item: amount * rates[item] for item, amount in current.items()}
    return baseline, tickets, current, rates, costs


def package_parts(path: Path) -> tuple[dict[str, bytes], etree._Element]:
    with zipfile.ZipFile(path) as package:
        damaged = package.testzip()
        if damaged:
            raise ValueError(f"CRC 损坏: {damaged}")
        parts = {name: package.read(name) for name in package.namelist()}
    return parts, etree.fromstring(parts["word/document.xml"])


def body_text(root: etree._Element) -> str:
    return "\n".join("".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS)) for paragraph in root.xpath("//w:body//w:p", namespaces=NS))


def table_by_header(document: Document, labels: list[str]):
    for table in document.tables:
        if table.rows and [cell.text.strip() for cell in table.rows[0].cells][:len(labels)] == labels:
            return table
    return None


def score_current(document: Document, current, rates, costs) -> tuple[float, str]:
    headers = ["Item", "Final amount (lb)", "Unit rate", "Extended cost", "Change origin", "Evidence / control note"]
    table = table_by_header(document, headers)
    if table is None or len(table.rows) != len(current) + 2:
        return 0.0, f"当前表行数异常: {len(table.rows) if table else 0}"
    origins = {"apple": "T01", "cabbage": "Baseline", "yogurt": "T03"}
    source_checks = {"apple": ["T01", "row 2"], "cabbage": ["baseline", "row 3"], "yogurt": ["T03", "evaluation"]}
    hits = 0
    total = len(current) * 6 + 6
    details = []
    for row, (item, amount) in zip(table.rows[1:-1], current.items()):
        cells = [cell.text.strip() for cell in row.cells]
        checks = [
            cells[0] == item,
            cells[1] == number(amount),
            cells[2] == f"{money(rates[item])}/lb",
            cells[3] == money(costs[item]),
            cells[4] == origins[item],
            all(fragment.lower() in cells[5].lower() for fragment in source_checks[item]),
        ]
        hits += sum(checks)
        details.append(f"{item}={sum(checks)}/6")
    final_total = sum(current.values())
    total_cost = sum(costs.values())
    cells = [cell.text.strip() for cell in table.rows[-1].cells]
    total_checks = [cells[0] == "TOTAL", cells[1] == number(final_total), cells[2] == "—", cells[3] == money(total_cost), str(len(current)) in cells[4], "Formula" in cells[5]]
    hits += sum(total_checks)
    return hits / total, ", ".join(details) + f", total={sum(total_checks)}/6"


def score_changes(document: Document, tickets: list[Ticket]) -> tuple[float, str]:
    headers = ["Ticket", "Source file", "Item", "Action", "Before", "After", "Delta", "Disposition / source location"]
    table = table_by_header(document, headers)
    if table is None or len(table.rows) != len(tickets) + 1:
        return 0.0, f"变更台账行数异常: {len(table.rows) if table else 0}"
    hits = 0
    total = len(tickets) * 8
    details = []
    for row, ticket in zip(table.rows[1:], tickets):
        cells = [cell.text.strip() for cell in row.cells]
        checks = [cells[0] == ticket.identifier, cells[1] == ticket.filename, cells[2] == ticket.item, cells[3] == ticket.action, cells[4] == number(ticket.before), cells[5] == number(ticket.after), cells[6] == signed(ticket.delta), "Applied once" in cells[7] and ticket.source_location in cells[7]]
        hits += sum(checks)
        details.append(f"{ticket.identifier}={sum(checks)}/8")
    return hits / total, ", ".join(details)


def score_reconciliation(document: Document, baseline, current) -> tuple[float, str]:
    headers = ["Baseline row", "Item", "Baseline", "Final", "Delta", "State", "Controlling evidence"]
    table = table_by_header(document, headers)
    baseline_lookup = {item: (row, amount) for row, item, amount in baseline}
    all_items = [item for _, item, _ in baseline] + [item for item in current if item not in baseline_lookup]
    if table is None or len(table.rows) != len(all_items) + 1:
        return 0.0, f"对账表行数异常: {len(table.rows) if table else 0}"
    states = {"apple": "updated", "cabbage": "unchanged", "onion": "removed", "yogurt": "added"}
    evidence = {"apple": "T01", "cabbage": "baseline only", "onion": "T02", "yogurt": "T03"}
    hits = 0
    total = len(all_items) * 7
    details = []
    for row, item in zip(table.rows[1:], all_items):
        source_row, before = baseline_lookup.get(item, (None, Decimal("0")))
        after = current.get(item, Decimal("0"))
        cells = [cell.text.strip() for cell in row.cells]
        checks = [cells[0] == (str(source_row) if source_row else "—"), cells[1] == item, cells[2] == number(before), cells[3] == number(after), cells[4] == signed(after - before), cells[5].lower() == states[item], cells[6].lower() == evidence[item].lower()]
        hits += sum(checks)
        details.append(f"{item}={sum(checks)}/7")
    return hits / total, ", ".join(details)


def score_calculations(document: Document, text: str, baseline, current, costs) -> tuple[float, str]:
    table = table_by_header(document, ["Control", "Calculation", "Result"])
    if table is None:
        return 0.0, "未找到计算控制表"
    rows = {row.cells[0].text.strip(): [cell.text.strip() for cell in row.cells] for row in table.rows[1:]}
    baseline_total = sum(amount for _, _, amount in baseline)
    final_total = sum(current.values())
    average = final_total / len(current)
    total_cost = sum(costs.values())
    expected = {
        "Baseline total amount": ("2 + 2 + 2", number(baseline_total)),
        "Final total amount": ("1 + 2 + 2", number(final_total)),
        "Signed amount change": ("final 5 − baseline 6", signed(final_total - baseline_total)),
        "Average final amount": ("final amount 5 ÷ 3 current lines", f"{average:.2f} lb"),
        "Estimated cost": ("$2.00 + $3.00 + $3.00", money(total_cost)),
    }
    checks = []
    for label, (formula, value) in expected.items():
        cells = rows.get(label, [])
        checks.extend([len(cells) == 3 and cells[1] == formula, len(cells) == 3 and cells[2] == value])
    lower = text.lower()
    checks.extend(["controlled — ready for handoff" in lower, "average final amount 1.67" in lower, "estimated current-list total $8.00" in lower])
    return sum(checks) / len(checks), f"checks={checks}, labels={list(rows)}"


def score_sources(document: Document) -> tuple[float, str]:
    table = table_by_header(document, ["Frozen source", "Control role", "Verification location"])
    expected = [
        ("baseline_shopping_list.xlsx", "Starting item/amount register", "Sheet1 rows 2–4"),
        ("change_T01_reduce_apple.json", "Approved quantity reduction", "evaluation[0].args.matches[0]"),
        ("change_T02_remove_onion.json", "Approved removal", "evaluation[0].args.keywords"),
        ("change_T03_add_yogurt.json", "Approved addition", "evaluation[0].args.matches[0:2]"),
        ("pricing_policy.json", "Named exception and other-item rate", "task + evaluation entries"),
    ]
    if table is None or len(table.rows) != len(expected) + 1:
        return 0.0, f"来源注册行数异常: {len(table.rows) if table else 0}"
    hits = 0
    total = len(expected) * 3
    details = []
    for row, wanted in zip(table.rows[1:], expected):
        cells = [cell.text.strip() for cell in row.cells]
        checks = [cells[0] == wanted[0], cells[1] == wanted[1], wanted[2] in cells[2]]
        hits += sum(checks)
        details.append(f"{wanted[0]}={sum(checks)}/3")
    return hits / total, ", ".join(details)


def score_comments(parts: dict[str, bytes], root: etree._Element) -> tuple[float, str]:
    if "word/comments.xml" not in parts:
        return 0.0, "无 Word 原生批注部件"
    comments_root = etree.fromstring(parts["word/comments.xml"])
    notes = ["".join(node.text or "" for node in comment.xpath(".//w:t", namespaces=NS)) for comment in comments_root.xpath("//w:comment", namespaces=NS)]
    anchored = []
    for start in root.xpath("//w:commentRangeStart", namespaces=NS):
        paragraph = start.getparent()
        anchored.append("".join(node.text or "" for node in paragraph.xpath(".//w:t", namespaces=NS)))
    apple_anchor = any("T01 update: baseline amount halved" in value for value in anchored)
    yogurt_anchor = any("T03 addition: approved yogurt line" in value for value in anchored)
    apple_note = any(all(fragment in note for fragment in ["T01", "change_T01_reduce_apple.json", "evaluation[0].args.matches[0]", "Sheet1 row 2"]) for note in notes)
    yogurt_note = any(all(fragment in note for fragment in ["T03", "change_T03_add_yogurt.json", "evaluation[0].args.matches[0:2]", "appended"]) for note in notes)
    checks = [len(notes) == 2, apple_anchor, yogurt_anchor, apple_note, yogurt_note]
    return sum(checks) / len(checks), f"comments={len(notes)}, anchors={anchored}, checks={checks}"


def score_navigation(root: etree._Element) -> tuple[float, str]:
    bookmarks = {node.get(f"{{{WNS}}}name", "") for node in root.xpath("//w:bookmarkStart", namespaces=NS)}
    anchors = [node.get(f"{{{WNS}}}anchor", "") for node in root.xpath("//w:hyperlink[@w:anchor]", namespaces=NS)]
    required = {"executive_reconciliation", "current_purchase", "change_ledger", "baseline_reconciliation", "calculation_notes", "ticket_t01", "ticket_t02", "ticket_t03"}
    instructions = [node.text or "" for node in root.xpath("//w:instrText", namespaces=NS)] + list(root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
    checks = [required <= bookmarks, len(anchors) >= 9, all(anchor in bookmarks for anchor in anchors), "TOC" in " ".join(map(str, instructions)).upper()]
    return sum(checks) / len(checks), f"bookmarks={len(bookmarks)}, required={len(required)}, links={len(anchors)}, checks={checks}"


def score_structure(parts: dict[str, bytes], root: etree._Element, document: Document) -> tuple[float, str]:
    sections = root.xpath("//w:sectPr", namespaces=NS)
    orientations = [bool(section.xpath("./w:pgSz[@w:orient='landscape']", namespaces=NS)) for section in sections]
    headers = " ".join("".join(node.text or "" for node in etree.fromstring(payload).xpath("//w:t", namespaces=NS)) for name, payload in parts.items() if name.startswith("word/header") and name.endswith(".xml"))
    footer_fields = []
    for name, payload in parts.items():
        if name.startswith("word/footer") and name.endswith(".xml"):
            footer_root = etree.fromstring(payload)
            footer_fields += [node.text or "" for node in footer_root.xpath("//w:instrText", namespaces=NS)]
            footer_fields += list(footer_root.xpath("//w:fldSimple/@w:instr", namespaces=NS))
    field_text = " ".join(map(str, footer_fields)).upper()
    protected_headers = ["Item", "Ticket", "Baseline row", "Control"]
    table_checks = []
    for label in protected_headers:
        tables = root.xpath("//w:tbl[w:tr[1]//w:t[contains(., $label)]]", namespaces=NS, label=label)
        if not tables:
            table_checks.append(False)
            continue
        rows = tables[0].xpath("./w:tr", namespaces=NS)
        table_checks.append(bool(rows[0].xpath("./w:trPr/w:tblHeader", namespaces=NS)) and all(row.xpath("./w:trPr/w:cantSplit", namespaces=NS) for row in rows))
    properties = document.core_properties
    checks = [len(sections) == 3, orientations == [False, True, False], "READY FOR HANDOFF" in headers, "PAGE" in field_text and "NUMPAGES" in field_text, all(table_checks), properties.title == "Controlled Shopping-List Purchase Brief", properties.author == "Purchase Control Review"]
    return sum(checks) / len(checks), f"sections={len(sections)}, orientations={orientations}, header={headers!r}, footer={field_text!r}, tables={table_checks}, title={properties.title!r}, author={properties.author!r}"


def score_boundary(parts: dict[str, bytes], root: etree._Element, text: str, document: Document) -> tuple[float, str]:
    relationships = []
    for name, payload in parts.items():
        if name.endswith(".rels"):
            relationships.extend(etree.fromstring(payload).findall(f"{{{RELNS}}}Relationship"))
    external = any(relationship.get("TargetMode") == "External" for relationship in relationships)
    unsafe = any(name.endswith("vbaProject.bin") or "oleObject" in name or name.endswith("afchunk.mht") for name in parts)
    drawings = bool(root.xpath("//w:drawing | //w:pict", namespaces=NS))
    lower = text.lower()
    forbidden_patterns = [
        r"(?:store|vendor|brand)\s*[:=]\s*\w+",
        r"tax\s*(?:rate|=|:)\s*\d+",
        r"discount\s*(?:=|:)\s*\d+",
        r"budget\s*(?:=|:)\s*\$?\d+",
        r"purchase date\s*[:=]",
        r"price guarantee\s*[:=]",
    ]
    forbidden = [pattern for pattern in forbidden_patterns if re.search(pattern, lower)]
    explicit = all(fragment in lower for fragment in ["policy-based estimates", "not a store quotation", "no store", "vendor", "brand", "tax", "discount", "budget", "package size", "purchase date"])
    pricing = all(fragment in text for fragment in ["apple uses the named $2.00-per-pound exception", "$1.50-per-pound rate", "Amount is interpreted as pounds only"])
    native = len(document.tables) >= 7 and len(text) >= 1800
    checks = [not external, not unsafe, not drawings, not forbidden, explicit, pricing, native]
    return sum(checks) / len(checks), f"external={external}, unsafe={unsafe}, drawings={drawings}, forbidden={forbidden}, text_chars={len(text)}, tables={len(document.tables)}, checks={checks}"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--result", required=True, type=Path)
    parser.add_argument("--reward-file", required=True, type=Path)
    arguments = parser.parse_args()
    source_dir = Path(__file__).resolve().parents[1] / "data/input_files"
    output = arguments.artifact_dir / OUTPUT_NAME
    artifacts = [path for path in arguments.artifact_dir.iterdir() if not path.name.startswith("~$")] if arguments.artifact_dir.is_dir() else []
    docx_files = [path for path in artifacts if path.suffix.lower() == ".docx"]
    results = []
    try:
        baseline, tickets, current, rates, costs = derive_truth(source_dir)
        if len(artifacts) != 1 or len(docx_files) != 1 or docx_files[0].name != OUTPUT_NAME or not output.is_file() or not zipfile.is_zipfile(output):
            raise ValueError(f"交付文件={[path.name for path in artifacts]}")
        parts, root = package_parts(output)
        if not {"[Content_Types].xml", "word/document.xml", "word/styles.xml"} <= parts.keys():
            raise ValueError("缺少 DOCX 核心部件")
        document = Document(output)
        text = body_text(root)
        results.append(criterion("R001", 1.0, f"唯一交付={OUTPUT_NAME}，OOXML 可解析，members={len(parts)}"))
        score, evidence = score_current(document, current, rates, costs); results.append(criterion("R002", score, evidence))
        score, evidence = score_changes(document, tickets); results.append(criterion("R003", score, evidence))
        score, evidence = score_reconciliation(document, baseline, current); results.append(criterion("R004", score, evidence))
        score, evidence = score_calculations(document, text, baseline, current, costs); results.append(criterion("R005", score, evidence))
        score, evidence = score_sources(document); results.append(criterion("R006", score, evidence))
        score, evidence = score_comments(parts, root); results.append(criterion("R007", score, evidence))
        score, evidence = score_navigation(root); results.append(criterion("R008", score, evidence))
        score, evidence = score_structure(parts, root, document); results.append(criterion("R009", score, evidence))
        score, evidence = score_boundary(parts, root, text, document); results.append(criterion("R010", score, evidence))
    except Exception as exc:
        results = [criterion("R001", 0.0, f"读取或唯一文件门槛失败: {exc}")]
    values = {item["criterion_id"]: item["score"] for item in results}
    reward = sum(WEIGHTS[identifier] * values.get(identifier, 0.0) for identifier in WEIGHTS)
    hurdle_failures = []
    if values.get("R001", 0.0) < 1.0:
        hurdle_failures.append("R001")
        reward = 0.0
    caps = {"R002": .48, "R003": .62, "R004": .68, "R005": .72, "R006": .76, "R007": .78, "R009": .79, "R010": .75}
    for identifier, cap in caps.items():
        if values.get(identifier, 0.0) < 1.0:
            hurdle_failures.append(identifier)
            reward = min(reward, cap)
    reward = round(reward, 6)
    payload = {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [hashlib.sha256(output.read_bytes()).hexdigest()] if output.is_file() else [],
        "criterion_results": results,
        "reward": reward,
        "passed": reward >= .8 and not hurdle_failures,
        "hurdle_failures": hurdle_failures,
    }
    arguments.result.parent.mkdir(parents=True, exist_ok=True)
    arguments.result.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    arguments.reward_file.parent.mkdir(parents=True, exist_ok=True)
    arguments.reward_file.write_text(f"{reward:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if payload["passed"] else 1)


if __name__ == "__main__":
    main()
