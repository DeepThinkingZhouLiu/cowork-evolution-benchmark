#!/usr/bin/env python3
"""SpreadsheetBench-2 Debugging 09_01 的确定性 Harbor Judge。"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import warnings
import zipfile
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries


TASK_ID = "cowork-evo/spreadsheetbench2-09-01-lbo-formula-repair"
RUBRIC_VERSION = "spreadsheetbench2-09-01-lbo-formula-repair-v0.1"
JUDGE_VERSION = "1.0.0"
PASS_THRESHOLD = 0.8
OUTPUT_NAMES = ["repair_workbook.py", "09_01_repaired.xlsx", "repair_log.json"]
ANSWER_RANGES = {
    "LBO": "A2:X218",
    "Financials": "A2:Q29",
    "Operating Cases": "A2:Q47",
}
INPUT_SHA256 = "f8a833086c28d4fd21e1bd11a083e2e8bd022002f5e0f6c2c57ae4ff9df80a44"
UPSTREAM_GOLDEN_SHA256 = "16d82179dff96d5e35d0617cdc0a784fbe455062f78ed7c91a0421a65f88a83a"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}

CRITERION_SPECS = {
    "R001": (8, "hurdle", "三个指定交付物完整且可解析"),
    "R002": (13, "hurdle", "29 个目标公式全部正确"),
    "R003": (12, "hurdle", "答案区域无非目标回归"),
    "R004": (12, "hurdle", "processor 在原输入上 fresh replay 正确"),
    "R005": (15, "hurdle", "processor 通过新增同类错误和源值变化的变异回放"),
    "R006": (9, "hurdle", "工作簿结构和非目标样式完整保留"),
    "R007": (6, "hurdle", "目标仍为公式、无新 #REF! 且启用完整重算"),
    "R008": (7, "hurdle", "修复日志与实际 4 类、8 节点、29 单元格一致"),
    "R009": (7, "hurdle", "输入和上游溯源文件哈希正确"),
    "R010": (5, "positive", "LibreOffice 可打开并导出 PDF"),
    "R011": (4, "positive", "工作簿保持原生可编辑财务模型"),
    "R012": (-25, "penalty", "交付缺失、损坏或不可解析"),
    "R013": (-25, "penalty", "未修复原输入或目标公式错误"),
    "R014": (-20, "penalty", "非目标公式或常量发生回归"),
    "R015": (-25, "penalty", "processor 不可回放、复制固定答案或不能处理变异"),
    "R016": (-18, "penalty", "工作簿结构或样式损坏"),
    "R017": (-12, "penalty", "修复日志虚报"),
}

HURDLE_CAPS = {
    "R001": 0.0,
    "R002": 0.45,
    "R003": 0.55,
    "R004": 0.35,
    "R005": 0.40,
    "R006": 0.65,
    "R007": 0.70,
    "R008": 0.70,
    "R009": 0.50,
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def criterion(identifier: str, passed: bool, evidence: str) -> dict[str, Any]:
    weight, kind, description = CRITERION_SPECS[identifier]
    return {
        "criterion_id": identifier,
        "passed": passed,
        "score": 1.0 if passed else 0.0,
        "weight": weight,
        "type": kind,
        "description": description,
        "method": "deterministic",
        "evidence": evidence,
    }


def normalized_formula(value: Any) -> str | None:
    if not isinstance(value, str) or not value.startswith("="):
        return None
    normalized = value.replace("$", "").replace(" ", "").upper()
    if normalized.startswith("=+"):
        normalized = "=" + normalized[2:]
    return normalized


def logically_equal(left: Any, right: Any) -> bool:
    left_formula = normalized_formula(left)
    right_formula = normalized_formula(right)
    if left_formula is not None or right_formula is not None:
        return left_formula is not None and left_formula == right_formula
    if left in (None, "") and right in (None, ""):
        return True
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return abs(float(left) - float(right)) <= max(1e-10, 1e-10 * max(abs(float(left)), abs(float(right))))
    return left == right


def target_formulas(include_mutation: bool = False) -> dict[tuple[str, str], str]:
    result: dict[tuple[str, str], str] = {}
    for column_number in range(7, 15):
        column = get_column_letter(column_number)
        result[("LBO", f"{column}66")] = f"={column}64+{column}75"
        result[("LBO", f"{column}69")] = f"=SUM({column}66:{column}68)"
        result[("LBO", f"{column}175")] = f"={column}136+{column}137+{column}145+{column}152"
    start = 5 if include_mutation else 10
    for column_number in range(start, 15):
        column = get_column_letter(column_number)
        result[("Financials", f"{column}18")] = f"={column}12-SUM({column}15:{column}17)"
    return result


def iter_answer_cells(workbook: Any) -> Iterable[tuple[str, str, Any]]:
    for sheet_name, cell_range in ANSWER_RANGES.items():
        sheet = workbook[sheet_name]
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        for row in range(min_row, max_row + 1):
            for column in range(min_col, max_col + 1):
                cell = sheet.cell(row, column)
                yield sheet_name, cell.coordinate, cell


def load_formula_workbook(path: Path) -> Any:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return load_workbook(path, data_only=False, keep_links=True)


def inspect_targets(path: Path, include_mutation: bool = False) -> tuple[bool, list[str]]:
    expected = target_formulas(include_mutation)
    workbook = load_formula_workbook(path)
    errors = []
    try:
        for (sheet_name, coordinate), formula in expected.items():
            actual = workbook[sheet_name][coordinate].value
            if normalized_formula(actual) != normalized_formula(formula):
                errors.append(f"{sheet_name}!{coordinate}: {actual!r} != {formula!r}")
    finally:
        workbook.close()
    return not errors, errors


def inspect_regression(
    source_path: Path,
    output_path: Path,
    *,
    include_mutation: bool = False,
) -> tuple[bool, list[str]]:
    source = load_formula_workbook(source_path)
    output = load_formula_workbook(output_path)
    excluded = set(target_formulas(include_mutation))
    errors: list[str] = []
    try:
        if output.sheetnames != source.sheetnames:
            return False, [f"sheetnames: {output.sheetnames} != {source.sheetnames}"]
        for sheet_name, coordinate, source_cell in iter_answer_cells(source):
            if (sheet_name, coordinate) in excluded:
                continue
            output_cell = output[sheet_name][coordinate]
            if not logically_equal(source_cell.value, output_cell.value):
                errors.append(
                    f"{sheet_name}!{coordinate}: {output_cell.value!r} != input {source_cell.value!r}"
                )
                if len(errors) >= 40:
                    break
    finally:
        source.close()
        output.close()
    return not errors, errors


def defined_name_signature(workbook: Any) -> list[tuple[Any, ...]]:
    result = []
    for name, item in workbook.defined_names.items():
        result.append(
            (
                name,
                item.attr_text,
                item.localSheetId,
                item.hidden,
                item.function,
                item.vbProcedure,
            )
        )
    return sorted(result, key=lambda item: (str(item[0]), str(item[2])))


def inspect_structure_and_style(source_path: Path, output_path: Path) -> tuple[bool, list[str]]:
    source = load_formula_workbook(source_path)
    output = load_formula_workbook(output_path)
    errors: list[str] = []
    try:
        if output.sheetnames != source.sheetnames:
            errors.append(f"工作表顺序不同：{output.sheetnames}")
        if defined_name_signature(output) != defined_name_signature(source):
            errors.append("定义名称集合或定义内容发生变化")
        for sheet_name in source.sheetnames:
            source_sheet = source[sheet_name]
            output_sheet = output[sheet_name]
            if (output_sheet.max_row, output_sheet.max_column) != (
                source_sheet.max_row,
                source_sheet.max_column,
            ):
                errors.append(
                    f"{sheet_name} 尺寸变化：{output_sheet.max_row}x{output_sheet.max_column}"
                )
            if {str(item) for item in output_sheet.merged_cells.ranges} != {
                str(item) for item in source_sheet.merged_cells.ranges
            }:
                errors.append(f"{sheet_name} 合并单元格变化")
            if output_sheet.freeze_panes != source_sheet.freeze_panes:
                errors.append(f"{sheet_name} 冻结窗格变化")
        for sheet_name, coordinate, source_cell in iter_answer_cells(source):
            output_cell = output[sheet_name][coordinate]
            if source_cell._style != output_cell._style:
                errors.append(f"{sheet_name}!{coordinate} 样式变化")
                if len(errors) >= 40:
                    break
    finally:
        source.close()
        output.close()
    return not errors, errors


def inspect_formula_integrity(path: Path) -> tuple[bool, str]:
    workbook = load_formula_workbook(path)
    broken = []
    formula_count = 0
    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_count += 1
                        if "#REF!" in cell.value.upper():
                            broken.append(f"{sheet.title}!{cell.coordinate}")
        calculation = getattr(workbook, "calculation", None)
        calc_mode = getattr(calculation, "calcMode", None)
        full_calc = bool(getattr(calculation, "fullCalcOnLoad", False))
        force_calc = bool(getattr(calculation, "forceFullCalc", False))
    finally:
        workbook.close()
    ok = formula_count >= 900 and not broken and calc_mode == "auto" and full_calc and force_calc
    return ok, f"公式={formula_count}，#REF!={broken[:10]}，calcMode={calc_mode}，full={full_calc}，force={force_calc}"


def inspect_native_editability(source_path: Path, output_path: Path) -> tuple[bool, str]:
    source = load_formula_workbook(source_path)
    output = load_formula_workbook(output_path)
    try:
        source_formulas = sum(
            isinstance(cell.value, str) and cell.value.startswith("=")
            for sheet in source.worksheets
            for row in sheet.iter_rows()
            for cell in row
        )
        output_formulas = sum(
            isinstance(cell.value, str) and cell.value.startswith("=")
            for sheet in output.worksheets
            for row in sheet.iter_rows()
            for cell in row
        )
        ok = output.sheetnames == source.sheetnames and output_formulas == source_formulas
        return ok, f"输入公式={source_formulas}，输出公式={output_formulas}，工作表={output.sheetnames}"
    finally:
        source.close()
        output.close()


def inspect_log(log_path: Path, input_path: Path, output_path: Path, expected_cells: int) -> tuple[bool, str]:
    try:
        payload = json.loads(log_path.read_text(encoding="utf-8"))
    except Exception as error:
        return False, f"日志无法解析：{error}"
    expected_set = {f"{sheet}!{coordinate}" for sheet, coordinate in target_formulas(expected_cells == 34)}
    actual_set = set(payload.get("cells", []))
    checks = [
        payload.get("source_task") == "SpreadsheetBench-2/Debugging/09_01",
        payload.get("input_sha256") == sha256(input_path),
        payload.get("output_sha256") == sha256(output_path),
        payload.get("repair_groups") == 4,
        payload.get("repair_nodes") == (13 if expected_cells == 34 else 8),
        payload.get("repaired_formula_cells") == expected_cells,
        actual_set == expected_set,
        isinstance(payload.get("repairs"), list),
    ]
    return all(checks), f"checks={checks}，cells={len(actual_set)}，nodes={payload.get('repair_nodes')}"


def inspect_source_package(task_root: Path, input_path: Path) -> tuple[bool, str]:
    manifest_path = task_root / "metadata" / "input-manifest.json"
    upstream_golden = task_root / "adaptation" / "upstream-09_golden.xlsx"
    full_manifest = task_root / "adaptation" / "upstream-dataset.json"
    task_record = task_root / "adaptation" / "upstream-task-record.json"
    dataset_readme = task_root / "adaptation" / "upstream-dataset-README.md"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        records = json.loads(full_manifest.read_text(encoding="utf-8"))
        record = json.loads(task_record.read_text(encoding="utf-8"))
        selected = next(item for item in records if item.get("id") == "09_01")
        checks = [
            sha256(input_path) == INPUT_SHA256,
            input_path.stat().st_size == 58415,
            manifest["files"][0]["sha256"] == INPUT_SHA256,
            manifest["files"][0]["size_bytes"] == 58415,
            sha256(upstream_golden) == UPSTREAM_GOLDEN_SHA256,
            sha256(full_manifest) == "f4996f9f52bba2a84b7ca3e79eb865f334c92e149e5dae6ee61828ea8c1ff9cb",
            sha256(dataset_readme) == "fa2398b049fa96d0cc384cc365adc226cb18036388810f24e5c5fcd85cf2da70",
            selected == record,
        ]
        return all(checks), f"checks={checks}，input={sha256(input_path)}，upstream_golden={sha256(upstream_golden)}"
    except Exception as error:
        return False, f"溯源检查失败：{error}"


def sheet_paths(entries: dict[str, bytes]) -> dict[str, str]:
    workbook = ET.fromstring(entries["xl/workbook.xml"])
    relationships = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
    targets = {
        item.attrib["Id"]: item.attrib["Target"]
        for item in relationships.findall("pr:Relationship", XML_NS)
    }
    result = {}
    for sheet in workbook.findall("m:sheets/m:sheet", XML_NS):
        relationship_id = sheet.attrib[f"{{{REL_NS}}}id"]
        result[sheet.attrib["name"]] = posixpath.normpath(
            posixpath.join("xl", targets[relationship_id])
        )
    return result


def build_mutated_input(source_path: Path, target_path: Path) -> None:
    with zipfile.ZipFile(source_path) as archive:
        infos = archive.infolist()
        entries = {info.filename: archive.read(info.filename) for info in infos}
    path = sheet_paths(entries)["Financials"]
    root = ET.fromstring(entries[path])
    cells = {
        cell.attrib["r"]: cell
        for cell in root.findall(".//m:sheetData/m:row/m:c", XML_NS)
    }
    value = cells["G17"].find("m:v", XML_NS)
    if value is None:
        raise ValueError("无法定位变异单元格")
    value.text = "70.8"
    for column in "EFGHI":
        formula = cells[f"{column}18"].find("m:f", XML_NS)
        if formula is None or not formula.text:
            raise ValueError(f"无法定位变异公式 {column}18")
        formula.text = f"+{column}12-SUM({column}15:{column}17)-{column}17"
        cached = cells[f"{column}18"].find("m:v", XML_NS)
        if cached is not None:
            cells[f"{column}18"].remove(cached)
    entries[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(target_path, "w") as archive:
        for info in infos:
            archive.writestr(info, entries[info.filename])


def run_processor(processor: Path, input_path: Path, work_dir: Path) -> tuple[Path, Path, str, bool]:
    output_path = work_dir / "repaired.xlsx"
    log_path = work_dir / "repair.json"
    try:
        completed = subprocess.run(
            [
                sys.executable,
                str(processor),
                "--input",
                str(input_path),
                "--output",
                str(output_path),
                "--log",
                str(log_path),
            ],
            cwd=processor.parent,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=90,
            check=False,
        )
        success = completed.returncode == 0 and output_path.is_file() and log_path.is_file()
        return output_path, log_path, completed.stdout[-1200:], success
    except Exception as error:
        return output_path, log_path, str(error), False


def render_workbook(path: Path) -> tuple[bool, str]:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    pdfinfo = shutil.which("pdfinfo")
    if not libreoffice or not pdfinfo:
        return False, "缺少 LibreOffice 或 pdfinfo"
    with tempfile.TemporaryDirectory(prefix="cowork-evo-ssb2-render-") as directory:
        target = Path(directory)
        profile = target / "lo-profile"
        env = os.environ.copy()
        completed = subprocess.run(
            [
                libreoffice,
                "--headless",
                f"-env:UserInstallation={profile.resolve().as_uri()}",
                "--convert-to",
                "pdf",
                "--outdir",
                str(target),
                str(path),
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=150,
            check=False,
            env=env,
        )
        pdf = target / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf.is_file() or pdf.stat().st_size < 1000:
            return False, f"rc={completed.returncode}，pdf={pdf.is_file()}，日志={completed.stdout[-500:]}"
        info = subprocess.run(
            [pdfinfo, str(pdf)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        pages = int(match.group(1)) if match else 0
        return pages > 0, f"PDF={pdf.stat().st_size} bytes，pages={pages}，rc={completed.returncode}"


def build_result(results: list[dict[str, Any]], artifact_paths: list[Path]) -> dict[str, Any]:
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(
        item["weight"] for item in results if item["weight"] > 0 and item["passed"]
    )
    penalties = sum(
        item["weight"] for item in results if item["weight"] < 0 and item["passed"]
    )
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [
        item["criterion_id"]
        for item in results
        if item["type"] == "hurdle" and not item["passed"]
    ]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[item] for item in hurdle_failures))
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": score,
                "evidence": item["evidence"],
                "method": item["method"],
                "raw": {
                    "description": item["description"],
                    "weight": item["weight"],
                    "type": item["type"],
                    "passed": item["passed"],
                },
            }
        )
    digests = [sha256(path) for path in artifact_paths if path.is_file()]
    return {
        "task_id": TASK_ID,
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": digests or [hashlib.sha256(b"").hexdigest()],
        "criterion_results": criterion_results,
        "reward": round(reward, 6),
        "passed": reward >= PASS_THRESHOLD,
        "hurdle_failures": hurdle_failures,
    }


def evaluate(artifact_dir: Path, render: bool = True) -> dict[str, Any]:
    task_root = Path(__file__).resolve().parents[1]
    input_path = task_root / "data" / "input_files" / "09_01_input.xlsx"
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results: list[dict[str, Any]] = []

    missing = [name for name, path in paths.items() if not path.is_file() or path.stat().st_size == 0]
    integrity_ok = not missing
    if integrity_ok:
        try:
            with zipfile.ZipFile(paths["09_01_repaired.xlsx"]) as archive:
                integrity_ok = archive.testzip() is None
            json.loads(paths["repair_log.json"].read_text(encoding="utf-8"))
            compile(paths["repair_workbook.py"].read_text(encoding="utf-8"), str(paths["repair_workbook.py"]), "exec")
        except Exception as error:
            integrity_ok = False
            missing.append(f"解析失败：{error}")
    results.append(criterion("R001", integrity_ok, f"缺失或损坏={missing}"))

    if not integrity_ok:
        results.extend(
            [
                criterion("R002", False, "前置完整性失败"),
                criterion("R003", False, "前置完整性失败"),
                criterion("R004", False, "前置完整性失败"),
                criterion("R005", False, "前置完整性失败"),
                criterion("R006", False, "前置完整性失败"),
                criterion("R007", False, "前置完整性失败"),
                criterion("R008", False, "前置完整性失败"),
                criterion("R009", False, "前置完整性失败"),
                criterion("R010", False, "前置完整性失败"),
                criterion("R011", False, "前置完整性失败"),
                criterion("R012", True, f"缺失或损坏={missing}"),
                criterion("R013", True, "无法验证目标修复"),
                criterion("R014", True, "无法验证非目标回归"),
                criterion("R015", True, "无法执行 processor"),
                criterion("R016", True, "无法验证结构"),
                criterion("R017", True, "日志缺失或损坏"),
            ]
        )
        return build_result(results, list(paths.values()))

    target_ok, target_errors = inspect_targets(paths["09_01_repaired.xlsx"])
    results.append(criterion("R002", target_ok, f"错误数={len(target_errors)}，样例={target_errors[:8]}"))

    regression_ok, regression_errors = inspect_regression(input_path, paths["09_01_repaired.xlsx"])
    results.append(criterion("R003", regression_ok, f"回归数={len(regression_errors)}，样例={regression_errors[:8]}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-ssb2-replay-") as directory:
        replay_dir = Path(directory)
        replay_output, replay_log, replay_stdout, replay_ran = run_processor(
            paths["repair_workbook.py"], input_path, replay_dir / "canonical"
        )
        canonical_target = canonical_regression = canonical_log = False
        if replay_ran:
            canonical_target, _ = inspect_targets(replay_output)
            canonical_regression, _ = inspect_regression(input_path, replay_output)
            canonical_log, _ = inspect_log(replay_log, input_path, replay_output, 29)
        canonical_ok = replay_ran and canonical_target and canonical_regression and canonical_log
        results.append(
            criterion(
                "R004",
                canonical_ok,
                f"run={replay_ran}，target={canonical_target}，regression={canonical_regression}，log={canonical_log}，stdout={replay_stdout}",
            )
        )

        mutation_dir = replay_dir / "mutation"
        mutation_dir.mkdir(parents=True)
        mutated_input = mutation_dir / "09_01_mutated.xlsx"
        build_mutated_input(input_path, mutated_input)
        mutated_output, mutated_log, mutation_stdout, mutation_ran = run_processor(
            paths["repair_workbook.py"], mutated_input, mutation_dir / "run"
        )
        mutation_target = mutation_regression = mutation_log_ok = mutation_value = False
        if mutation_ran:
            mutation_target, _ = inspect_targets(mutated_output, include_mutation=True)
            mutation_regression, _ = inspect_regression(
                mutated_input, mutated_output, include_mutation=True
            )
            mutation_log_ok, _ = inspect_log(mutated_log, mutated_input, mutated_output, 34)
            workbook = load_formula_workbook(mutated_output)
            try:
                mutation_value = logically_equal(workbook["Financials"]["G17"].value, 70.8)
            finally:
                workbook.close()
        mutation_ok = (
            mutation_ran
            and mutation_target
            and mutation_regression
            and mutation_log_ok
            and mutation_value
        )
        results.append(
            criterion(
                "R005",
                mutation_ok,
                f"run={mutation_ran}，target34={mutation_target}，regression={mutation_regression}，log={mutation_log_ok}，G17=70.8={mutation_value}，stdout={mutation_stdout}",
            )
        )

    structure_ok, structure_errors = inspect_structure_and_style(
        input_path, paths["09_01_repaired.xlsx"]
    )
    results.append(criterion("R006", structure_ok, f"问题数={len(structure_errors)}，样例={structure_errors[:8]}"))

    formula_integrity, formula_evidence = inspect_formula_integrity(paths["09_01_repaired.xlsx"])
    results.append(criterion("R007", formula_integrity, formula_evidence))

    log_ok, log_evidence = inspect_log(
        paths["repair_log.json"], input_path, paths["09_01_repaired.xlsx"], 29
    )
    results.append(criterion("R008", log_ok, log_evidence))

    source_ok, source_evidence = inspect_source_package(task_root, input_path)
    results.append(criterion("R009", source_ok, source_evidence))

    render_ok, render_evidence = (render_workbook(paths["09_01_repaired.xlsx"]) if render else (True, "测试模式跳过渲染"))
    results.append(criterion("R010", render_ok, render_evidence))

    native_ok, native_evidence = inspect_native_editability(
        input_path, paths["09_01_repaired.xlsx"]
    )
    results.append(criterion("R011", native_ok, native_evidence))

    unchanged = sha256(paths["09_01_repaired.xlsx"]) == sha256(input_path)
    results.extend(
        [
            criterion("R012", False, "交付完整可解析，未触发"),
            criterion("R013", unchanged or not target_ok, f"unchanged={unchanged}，target_ok={target_ok}"),
            criterion("R014", not regression_ok, f"regression_ok={regression_ok}"),
            criterion("R015", not (canonical_ok and mutation_ok), f"canonical={canonical_ok}，mutation={mutation_ok}"),
            criterion("R016", not structure_ok, f"structure_ok={structure_ok}"),
            criterion("R017", not log_ok, f"log_ok={log_ok}"),
        ]
    )
    return build_result(results, list(paths.values()))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", required=True, type=Path)
    parser.add_argument("--skip-render", action="store_true")
    parser.add_argument("--json-output", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    result = evaluate(args.artifact_dir.resolve(), render=not args.skip_render)
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.json_output:
        args.json_output.write_text(payload, encoding="utf-8")
    print(payload, end="")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
