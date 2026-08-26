#!/usr/bin/env python3
"""逐题专用修复器：只修复 SpreadsheetBench-2 Debugging 09_01 的四组公式错误。"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}
CELL_RE = re.compile(r"^([A-Z]+)([0-9]+)$")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def column_number(name: str) -> int:
    value = 0
    for char in name:
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def column_name(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def split_coordinate(coordinate: str) -> tuple[str, int]:
    match = CELL_RE.fullmatch(coordinate)
    if not match:
        raise ValueError(f"无效单元格坐标：{coordinate}")
    return match.group(1), int(match.group(2))


def expand_range(cell_range: str) -> list[str]:
    if ":" not in cell_range:
        return [cell_range]
    start, end = cell_range.split(":", 1)
    start_col, start_row = split_coordinate(start)
    end_col, end_row = split_coordinate(end)
    return [
        f"{column_name(col)}{row}"
        for row in range(start_row, end_row + 1)
        for col in range(column_number(start_col), column_number(end_col) + 1)
    ]


def normalize_formula(formula: str) -> str:
    normalized = formula.replace("$", "").replace(" ", "").upper()
    if normalized.startswith("=+"):
        normalized = normalized[1:]
    elif normalized.startswith("="):
        normalized = normalized[1:]
    return normalized


def read_shared_strings(entries: dict[str, bytes]) -> list[str]:
    raw = entries.get("xl/sharedStrings.xml")
    if raw is None:
        return []
    root = ET.fromstring(raw)
    return ["".join(node.text or "" for node in item.findall(".//m:t", NS)) for item in root.findall("m:si", NS)]


def sheet_paths(entries: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(entries["xl/workbook.xml"])
    relationships = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall("pr:Relationship", NS)
    }
    result: dict[str, str] = {}
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = targets[relationship_id]
        result[sheet.attrib["name"]] = posixpath.normpath(posixpath.join("xl", target))
    return result


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str | None:
    cell_type = cell.attrib.get("t")
    if cell_type == "s":
        value = cell.find("m:v", NS)
        return shared_strings[int(value.text)] if value is not None and value.text is not None else None
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//m:t", NS))
    value = cell.find("m:v", NS)
    return value.text if value is not None else None


def cells_by_coordinate(root: ET.Element) -> dict[str, ET.Element]:
    return {cell.attrib["r"]: cell for cell in root.findall(".//m:sheetData/m:row/m:c", NS)}


def find_label_row(
    cells: dict[str, ET.Element],
    shared_strings: list[str],
    label: str,
    *,
    min_row: int | None = None,
    max_row: int | None = None,
) -> int:
    matches = []
    for coordinate, cell in cells.items():
        column, row = split_coordinate(coordinate)
        in_range = (min_row is None or row >= min_row) and (max_row is None or row <= max_row)
        if column == "B" and in_range and cell_text(cell, shared_strings) == label:
            matches.append(row)
    if len(matches) != 1:
        raise ValueError(f"标签 {label!r} 预期唯一，实际位于 {matches}")
    return matches[0]


def formula_groups_on_row(
    cells: dict[str, ET.Element], row: int, start_col: str, end_col: str
) -> list[tuple[str, ET.Element, ET.Element, list[str]]]:
    groups = []
    for col_no in range(column_number(start_col), column_number(end_col) + 1):
        coordinate = f"{column_name(col_no)}{row}"
        cell = cells.get(coordinate)
        if cell is None:
            continue
        formula = cell.find("m:f", NS)
        if formula is None or not formula.text:
            continue
        affected = expand_range(formula.attrib.get("ref", coordinate))
        groups.append((coordinate, cell, formula, affected))
    return groups


def clear_cached_values(cells: dict[str, ET.Element], coordinates: list[str]) -> None:
    for coordinate in coordinates:
        cell = cells.get(coordinate)
        if cell is None:
            continue
        cached = cell.find("m:v", NS)
        if cached is not None:
            cell.remove(cached)


def repair_formula_group(
    *,
    sheet: str,
    cells: dict[str, ET.Element],
    master_coordinate: str,
    formula: ET.Element,
    affected: list[str],
    desired: str,
    issue: str,
    repairs: list[dict[str, object]],
) -> None:
    before = formula.text or ""
    if normalize_formula(before) == normalize_formula(desired):
        return
    formula.text = desired
    clear_cached_values(cells, affected)
    repairs.append(
        {
            "issue": issue,
            "sheet": sheet,
            "formula_anchor": master_coordinate,
            "affected_cells": affected,
            "before": before,
            "after": desired,
        }
    )


def repair_lbo(root: ET.Element, shared_strings: list[str]) -> list[dict[str, object]]:
    cells = cells_by_coordinate(root)
    repairs: list[dict[str, object]] = []

    core_ebit = find_label_row(cells, shared_strings, "Core EBIT")
    ebit = find_label_row(cells, shared_strings, "EBIT")
    tax_step_up = find_label_row(cells, shared_strings, "(-) Tax Step-Up Amort")
    less_taxes = find_label_row(cells, shared_strings, "Less: Taxes")
    total_interest = find_label_row(cells, shared_strings, "Total Interest")
    fee = find_label_row(cells, shared_strings, "Fee on Unused Balance")

    interest_rows = []
    for coordinate, cell in cells.items():
        column, row = split_coordinate(coordinate)
        if column == "B" and cell_text(cell, shared_strings) == "Interest":
            interest_rows.append(row)
    interest_rows = sorted(interest_rows)
    if len(interest_rows) != 3:
        raise ValueError(f"预期找到三条债务 Interest 行，实际为 {interest_rows}")

    for master, _cell, formula, affected in formula_groups_on_row(cells, ebit, "G", "N"):
        col, _ = split_coordinate(master)
        repair_formula_group(
            sheet="LBO",
            cells=cells,
            master_coordinate=master,
            formula=formula,
            affected=affected,
            desired=f"+{col}{core_ebit}+{col}{tax_step_up}",
            issue="EBIT 错误引用 Adj. EBITDA，改为 Core EBIT 加税务摊销",
            repairs=repairs,
        )

    net_income = find_label_row(
        cells,
        shared_strings,
        "Net Income",
        min_row=ebit,
        max_row=less_taxes + 2,
    )
    for master, _cell, formula, affected in formula_groups_on_row(cells, net_income, "G", "N"):
        col, _ = split_coordinate(master)
        repair_formula_group(
            sheet="LBO",
            cells=cells,
            master_coordinate=master,
            formula=formula,
            affected=affected,
            desired=f"+SUM({col}{ebit}:{col}{less_taxes})",
            issue="Net Income 汇总范围重复包含 Core EBIT",
            repairs=repairs,
        )

    for master, _cell, formula, affected in formula_groups_on_row(cells, total_interest, "G", "N"):
        col, _ = split_coordinate(master)
        terms = [f"{col}{interest_rows[0]}", f"{col}{fee}", f"{col}{interest_rows[1]}", f"{col}{interest_rows[2]}"]
        repair_formula_group(
            sheet="LBO",
            cells=cells,
            master_coordinate=master,
            formula=formula,
            affected=affected,
            desired="+" + "+".join(terms),
            issue="Total Interest 提前计入 Cash Interest，导致后续再次扣减",
            repairs=repairs,
        )

    return repairs


def repair_financials(root: ET.Element, shared_strings: list[str]) -> list[dict[str, object]]:
    cells = cells_by_coordinate(root)
    repairs: list[dict[str, object]] = []
    revenue = find_label_row(cells, shared_strings, "Total Net Revenue")
    first_cost = find_label_row(cells, shared_strings, "Cost of goods sold")
    last_cost = find_label_row(cells, shared_strings, "General and administrative expenses")
    adjusted_ebitda = find_label_row(cells, shared_strings, "Adj. EBITDA")

    for master, _cell, formula, affected in formula_groups_on_row(cells, adjusted_ebitda, "E", "N"):
        col, _ = split_coordinate(master)
        repair_formula_group(
            sheet="Financials",
            cells=cells,
            master_coordinate=master,
            formula=formula,
            affected=affected,
            desired=f"+{col}{revenue}-SUM({col}{first_cost}:{col}{last_cost})",
            issue="Adj. EBITDA 在 SUM 后又重复扣除 G&A",
            repairs=repairs,
        )
    return repairs


def enable_full_recalculation(workbook_xml: bytes) -> bytes:
    root = ET.fromstring(workbook_xml)
    calc_pr = root.find("m:calcPr", NS)
    if calc_pr is None:
        calc_pr = ET.SubElement(root, f"{{{MAIN_NS}}}calcPr")
    calc_pr.set("calcMode", "auto")
    calc_pr.set("fullCalcOnLoad", "1")
    calc_pr.set("forceFullCalc", "1")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def repair(input_path: Path, output_path: Path, log_path: Path) -> dict[str, object]:
    if not input_path.is_file():
        raise FileNotFoundError(input_path)
    with zipfile.ZipFile(input_path) as archive:
        infos = archive.infolist()
        entries = {info.filename: archive.read(info.filename) for info in infos}

    shared_strings = read_shared_strings(entries)
    paths = sheet_paths(entries)
    for required in ("LBO", "Financials", "Operating Cases"):
        if required not in paths:
            raise ValueError(f"缺少工作表 {required}")

    lbo_root = ET.fromstring(entries[paths["LBO"]])
    financials_root = ET.fromstring(entries[paths["Financials"]])
    repairs = repair_lbo(lbo_root, shared_strings) + repair_financials(financials_root, shared_strings)
    entries[paths["LBO"]] = ET.tostring(lbo_root, encoding="utf-8", xml_declaration=True)
    entries[paths["Financials"]] = ET.tostring(financials_root, encoding="utf-8", xml_declaration=True)
    entries["xl/workbook.xml"] = enable_full_recalculation(entries["xl/workbook.xml"])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w") as archive:
        for info in infos:
            archive.writestr(info, entries[info.filename])

    repaired_cells = [cell for item in repairs for cell in item["affected_cells"]]
    issue_families = sorted({str(item["issue"]) for item in repairs})
    payload: dict[str, object] = {
        "source_task": "SpreadsheetBench-2/Debugging/09_01",
        "input_sha256": sha256(input_path),
        "output_sha256": sha256(output_path),
        "repair_groups": len(issue_families),
        "repair_nodes": len(repairs),
        "repaired_formula_cells": len(repaired_cells),
        "cells": [f"{item['sheet']}!{cell}" for item in repairs for cell in item["affected_cells"]],
        "issue_families": issue_families,
        "repairs": repairs,
        "preservation_method": "直接修改 OOXML 共享公式主节点；不重建工作簿、工作表、样式或业务数据",
    }
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return payload


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--log", required=True, type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    payload = repair(args.input.resolve(), args.output.resolve(), args.log.resolve())
    print(json.dumps({"repair_groups": payload["repair_groups"], "repaired_formula_cells": payload["repaired_formula_cells"]}, ensure_ascii=False))


if __name__ == "__main__":
    main()
