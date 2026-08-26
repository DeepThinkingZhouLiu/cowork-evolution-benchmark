#!/usr/bin/env python3
"""officeval_092 的 Harbor Judge。

基础单元格检查复用并保留上游 Apache-2.0 verifier；本文件补上上游
实现未覆盖的打印分页、截图替换、行高和原始辅助表保留检查。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import warnings
from pathlib import Path

from openpyxl import load_workbook

import upstream_verifier


OUTPUT_NAME = "智慧农业示范园建设项目报价表_完成版.xlsx"
MAIN_SHEET = "智慧农业报价表"
RUBRIC_VERSION = "omegause-officeval-092-adaptation-v0.1"
JUDGE_VERSION = "officeval-092-adaptation-v0.1"
EXPECTED_BREAKS = {23, 32, 39, 47, 58, 68, 78, 88, 98, 108, 118}
EXPECTED_FEES = [
    ["类别", "说明", "是否计入报价汇总"],
    ["生产温室区", "主体温室、育苗区、栽培区", "是"],
    ["环控系统", "通风、遮阳、湿帘风机", "是"],
    ["水肥系统", "水肥一体化及灌溉管线", "是"],
    ["物联网系统", "传感器、网关、平台", "是"],
    ["供配电", "照明、动力配电、桥架", "是"],
    ["分拣包装", "输送、称重、贴标", "是"],
    ["冷链仓储", "预冷库与周转库", "是"],
    ["运维配套", "家具、工具、标识", "是"],
]
TABLE_TITLE_PAIRS = [
    ("一、项目报价汇总", "含税总价"),
    ("1. 温室主体与基础", "1. 温室主体与基础 小计"),
    ("2. 地面、排水与道路", "2. 地面、排水与道路 小计"),
    ("3. 环控、水肥与电气", "3. 环控、水肥与电气 小计"),
    ("4. 物联网与分拣配套", "温室工程量清单合计"),
    ("四、温室主材预算清单", "四、温室主材预算清单 小计"),
    ("五、环控与照明物资清单", "五、环控与照明物资清单 小计"),
    ("六、物联网设备预算清单", "六、物联网设备预算清单 小计"),
    ("七、分拣包装与冷链设备清单", "七、分拣包装与冷链设备清单 小计"),
    ("八、运维家具与工具清单", "八、运维家具与工具清单 小计"),
    ("九、不锈钢玻璃与五金辅材清单", "九、不锈钢玻璃与五金辅材清单 小计"),
    ("十、清单总计与核对", "用于复核报价汇总与清单测算"),
]

CRITERION_SPECS = {
    "R001": ("交付文件存在、扩展名正确且可以打开", 2, "hurdle"),
    "R002": ("打印预览中每张完整表格位于同一页", 5, "hurdle"),
    "R003": ("表内字体、字号、颜色、垂直对齐和自动换行符合要求", 3, "positive"),
    "R004": ("表内文字完整显示，相似表格行高一致且不存在明显过窄行", 5, "positive"),
    "R005": ("A69:J76 环控与照明物资表内容、公式和可编辑性完整", 5, "positive"),
    "R006": ("A79:J86 物联网设备表内容、公式和可编辑性完整", 5, "positive"),
    "R007": ("工作表使用 A4 纵向、按一页宽缩放，分页仅位于表格间隔", 5, "positive"),
    "R008": ("F122 未把两张新增清单计入物资合计", -5, "penalty"),
    "R009": ("“费用类别”工作表或原始内容被破坏", -3, "penalty"),
    "R010": ("原截图已删除，新增区域全部由可编辑单元格组成", 5, "positive"),
}


def criterion(identifier: str, description: str, weight: int, passed: bool, detail: str = "", *, kind: str = "positive") -> dict:
    return {
        "criterion_id": identifier,
        "description": description,
        "weight": weight,
        "type": kind,
        "passed": bool(passed),
        "score": 1.0 if passed else 0.0,
        "detail": detail,
        "method": "deterministic",
    }


def render_page_pair_check(path: Path) -> tuple[bool, str]:
    """用真实 LibreOffice 打印渲染确认完整表格不跨页。"""

    libreoffice = shutil.which("libreoffice")
    pdftotext = shutil.which("pdftotext")
    if not libreoffice or not pdftotext:
        return False, "Judge 环境缺少 libreoffice 或 pdftotext"

    with tempfile.TemporaryDirectory(prefix="cowork-evo-092-") as temp_dir:
        completed = subprocess.run(
            [libreoffice, "--headless", "--convert-to", "pdf", "--outdir", temp_dir, str(path)],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        pdf_path = Path(temp_dir) / f"{path.stem}.pdf"
        if completed.returncode != 0 or not pdf_path.is_file():
            detail = (completed.stderr or completed.stdout or "转换失败").strip()
            return False, detail[-800:]
        text_result = subprocess.run(
            [pdftotext, "-layout", str(pdf_path), "-"],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
        if text_result.returncode != 0:
            return False, (text_result.stderr or "PDF 文本抽取失败")[-800:]
        pages = [page.replace(" ", "") for page in text_result.stdout.split("\f") if page.strip()]

    failures = []
    overflow_pages = [index + 1 for index, page in enumerate(pages) if "###" in page]
    if overflow_pages:
        failures.append(f"打印预览存在 ### 数值溢出：页码={overflow_pages}")
    for title, end_marker in TABLE_TITLE_PAIRS:
        matching_pages = [index + 1 for index, page in enumerate(pages) if title.replace(" ", "") in page]
        complete_pages = [
            index + 1
            for index, page in enumerate(pages)
            if title.replace(" ", "") in page and end_marker.replace(" ", "") in page
        ]
        if len(matching_pages) != 1 or len(complete_pages) != 1:
            failures.append(f"{title}: 标题页={matching_pages}，完整页={complete_pages}")
    return not failures, f"渲染页数={len(pages)}" if not failures else "；".join(failures)


def normalized_formula(value) -> str:
    return re.sub(r"\s+|\$", "", str(value or "")).upper()


def evaluate(artifact_dir: Path) -> dict:
    output_path = artifact_dir / OUTPUT_NAME
    results = []

    if not output_path.is_file():
        results.append(criterion("R001", "交付文件存在、扩展名正确且可以打开", 2, False, f"缺少 {OUTPUT_NAME}", kind="hurdle"))
        return build_result(results, output_path)

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            workbook = load_workbook(output_path, data_only=False)
            workbook_values = load_workbook(output_path, data_only=True)
    except Exception as exc:
        results.append(criterion("R001", "交付文件存在、扩展名正确且可以打开", 2, False, str(exc), kind="hurdle"))
        return build_result(results, output_path)

    main_exists = MAIN_SHEET in workbook.sheetnames
    results.append(criterion("R001", "交付文件存在、扩展名正确且可以打开", 2, main_exists, f"工作表={workbook.sheetnames}", kind="hurdle"))
    if not main_exists:
        return build_result(results, output_path)

    worksheet = workbook[MAIN_SHEET]
    worksheet_values = workbook_values[MAIN_SHEET]
    upstream_points = upstream_verifier.check_dimension2(output_path, workbook, workbook_values)
    upstream_by_name = {point.name: point for point in upstream_points}

    print_render_ok, print_render_detail = render_page_pair_check(output_path)
    break_ids = {int(item.id) for item in worksheet.row_breaks.brk}
    print_setup_ok = (
        worksheet.page_setup.orientation == "portrait"
        and str(worksheet.page_setup.paperSize) in {"9", "A4"}
        and worksheet.page_setup.fitToWidth == 1
        and worksheet.page_setup.fitToHeight in {0, None}
        and EXPECTED_BREAKS.issubset(break_ids)
    )
    results.append(
        criterion(
            "R002",
            "打印预览中每张完整表格位于同一页",
            5,
            print_render_ok and print_setup_ok,
            f"{print_render_detail}；分页={sorted(break_ids)}；A4纵向一页宽={print_setup_ok}",
            kind="hurdle",
        )
    )

    font_point = upstream_by_name["表内字体为微软雅黑10磅、黑色或深灰色，垂直居中并自动换行"]
    results.append(criterion("R003", font_point.name, 3, font_point.hit, font_point.detail))

    height_issues = upstream_verifier.row_height_truncation_issues(worksheet, 1, 124)
    results.append(
        criterion(
            "R004",
            "表内文字完整显示，相似表格行高一致且不存在明显过窄行",
            5,
            not height_issues,
            "全部符合" if not height_issues else "；".join(height_issues[:10]),
        )
    )

    env_point = upstream_by_name["A69:J76环控与照明物资表内容、公式和可编辑性完整"]
    results.append(criterion("R005", env_point.name, 5, env_point.hit, env_point.detail))

    iot_point = upstream_by_name["A79:J86物联网设备表内容、公式和可编辑性完整"]
    results.append(criterion("R006", iot_point.name, 5, iot_point.hit, iot_point.detail))

    results.append(
        criterion(
            "R007",
            "工作表使用 A4 纵向、按一页宽缩放，分页仅位于表格间隔",
            5,
            print_setup_ok,
            f"orientation={worksheet.page_setup.orientation}, paper={worksheet.page_setup.paperSize}, fit={worksheet.page_setup.fitToWidth}x{worksheet.page_setup.fitToHeight}, breaks={sorted(break_ids)}",
        )
    )

    formula_f122 = normalized_formula(worksheet["F122"].value)
    f122_ok = "G76" in formula_f122 and "G86" in formula_f122
    results.append(
        criterion(
            "R008",
            "F122 必须把 G76 与 G86 两张新增清单计入物资合计",
            -5,
            not f122_ok,
            f"F122={worksheet['F122'].value!r}",
            kind="penalty",
        )
    )

    fees_ok = "费用类别" in workbook.sheetnames
    fee_errors = []
    if fees_ok:
        fees_sheet = workbook["费用类别"]
        for row, expected_row in enumerate(EXPECTED_FEES, 1):
            actual_row = [fees_sheet.cell(row, column).value for column in range(1, 4)]
            if actual_row != expected_row:
                fee_errors.append(f"第{row}行期望={expected_row!r}，实际={actual_row!r}")
    results.append(
        criterion(
            "R009",
            "“费用类别”工作表及 A1:C9 原始内容必须完整保留",
            -3,
            not fees_ok or bool(fee_errors),
            "完整保留" if fees_ok and not fee_errors else "；".join(fee_errors[:5]) or "缺少工作表",
            kind="penalty",
        )
    )

    object_summary = upstream_verifier.workbook_object_summary(output_path, workbook)
    screenshot_removed = object_summary["media"] == 0 and object_summary["images_openpyxl"] == 0
    results.append(
        criterion(
            "R010",
            "原截图已被删除，新增区域全部由可编辑单元格组成",
            5,
            screenshot_removed,
            json.dumps(object_summary, ensure_ascii=False),
        )
    )

    # 额外验证公式可由上游逐题解析器得到准确数值。
    subtotal_values = {
        "G76": upstream_verifier.formula_result(worksheet, worksheet_values, "G76"),
        "G86": upstream_verifier.formula_result(worksheet, worksheet_values, "G86"),
    }
    if subtotal_values != {"G76": 448580.0, "G86": 90440.0}:
        for item in results:
            if item["criterion_id"] in {"R005", "R006"}:
                item["passed"] = False
                item["score"] = 0.0
                item["detail"] += f"；小计={subtotal_values}"

    return build_result(results, output_path)


def build_result(results: list[dict], artifact_path: Path | None = None) -> dict:
    """计算得分并输出符合仓库 judge-result schema 的结果。"""

    present_ids = {item["criterion_id"] for item in results}
    for identifier, (description, weight, kind) in CRITERION_SPECS.items():
        if identifier not in present_ids:
            results.append(
                criterion(
                    identifier,
                    description,
                    weight,
                    False,
                    "因交付文件不存在或无法打开，本判据未执行",
                    kind=kind,
                )
            )
    results.sort(key=lambda item: item["criterion_id"])

    positive_weight = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_weight)) if positive_weight else 0.0

    hurdle_failures = [item for item in results if item["type"] == "hurdle" and not item["passed"]]
    if any(item["criterion_id"] == "R001" for item in hurdle_failures):
        reward = 0.0
    elif hurdle_failures:
        reward = min(reward, 0.5)

    if artifact_path is not None and artifact_path.is_file():
        artifact_digest = hashlib.sha256(artifact_path.read_bytes()).hexdigest()
    else:
        # schema 要求至少一个摘要；空字节摘要明确表示没有可读取的交付文件。
        artifact_digest = hashlib.sha256(b"").hexdigest()

    normalized_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        normalized_results.append(
            {
                "criterion_id": item["criterion_id"],
                "score": score,
                "evidence": item["detail"],
                "method": item["method"],
                "raw": {
                    "description": item["description"],
                    "weight": item["weight"],
                    "type": item["type"],
                    "passed": item["passed"],
                },
            }
        )

    return {
        "task_id": "cowork-evo/officeval-092-screenshot-to-editable-table",
        "rubric_version": RUBRIC_VERSION,
        "judge_version": JUDGE_VERSION,
        "artifact_sha256": [artifact_digest],
        "reward": round(reward, 6),
        "passed": reward >= 0.7,
        "criterion_results": normalized_results,
        "hurdle_failures": [item["criterion_id"] for item in hurdle_failures],
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
