#!/usr/bin/env python3
"""单题回放、内容变异回放和针对性负例。

全部临时变异都发生在 TemporaryDirectory 中，不改写题包输入或 Golden。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.util import Inches, Pt


def load_judge(task_root: Path):
    spec = importlib.util.spec_from_file_location("budget_task_judge", task_root / "tests" / "judge.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("无法加载 Judge")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_solver(task_root: Path, input_dir: Path, output_dir: Path, workspace_dir: Path) -> None:
    completed = subprocess.run(
        [
            sys.executable,
            str(task_root / "solution" / "solve.py"),
            "--input-dir",
            str(input_dir),
            "--output-dir",
            str(output_dir),
            "--workspace-dir",
            str(workspace_dir),
        ],
        capture_output=True,
        text=True,
        timeout=360,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError((completed.stderr or completed.stdout or "solve.py 失败")[-3000:])


def copy_outputs(source: Path, target: Path) -> None:
    target.mkdir(parents=True, exist_ok=True)
    for name in ("budget_vs_actuals.xlsx", "budget_vs_actuals.pptx"):
        shutil.copy2(source / name, target / name)


def mutate_rd_salary(source: Path, target: Path) -> dict[str, str]:
    lines = source.read_text(encoding="utf-8").splitlines(keepends=True)
    columns = [part.strip('"') for part in lines[0].split("(", 1)[1].split(") FROM stdin", 1)[0].split(", ")]
    salary_index = columns.index("SALARY")
    department_index = columns.index("DEPARTMENT")
    employee_id_index = columns.index("EMPLOYEE_ID")
    for index in range(1, len(lines) - 1):
        suffix = "\n" if lines[index].endswith("\n") else ""
        fields = lines[index].rstrip("\r\n").split("\t")
        if fields[department_index] != "НИОКР":
            continue
        old_salary = Decimal(fields[salary_index])
        new_salary = old_salary + Decimal("400000000.00")
        fields[salary_index] = f"{new_salary:.2f}"
        lines[index] = "\t".join(fields) + suffix
        target.write_text("".join(lines), encoding="utf-8")
        return {
            "employee_id": fields[employee_id_index],
            "department": fields[department_index],
            "old_salary": f"{old_salary:.2f}",
            "new_salary": f"{new_salary:.2f}",
            "delta": "400000000.00",
        }
    raise AssertionError("没有找到可变异的 НИОКР 员工行")


def expect_rejected(judge, artifact_dir: Path, data_dir: Path, expected_failure: str) -> dict:
    result = judge.evaluate(artifact_dir, data_dir)
    if result["passed"] or result["score"] >= 0.85:
        raise AssertionError(f"负例未被拒绝：{artifact_dir} -> {result['score']}")
    if expected_failure not in result["hurdle_failures"]:
        raise AssertionError(
            f"负例未命中期望门槛 {expected_failure}：{result['hurdle_failures']}"
        )
    return result


def change_excel_actual(path: Path) -> None:
    workbook = load_workbook(path)
    sheet = workbook["Department Details"]
    sheet["C2"] = Decimal(str(sheet["C2"].value)) + Decimal("12345.00")
    workbook.save(path)
    workbook.close()


def break_excel_formula(path: Path) -> None:
    workbook = load_workbook(path)
    sheet = workbook["Department Details"]
    sheet["D2"] = "=1"
    workbook.save(path)
    workbook.close()


def change_ppt_actual(path: Path) -> None:
    presentation = Presentation(path)
    old = "$418,604,451.00"
    new = "$418,604,452.00"
    changed = False
    for shape in presentation.slides[2].shapes:
        if getattr(shape, "has_text_frame", False) and old in shape.text:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if old in run.text:
                        run.text = run.text.replace(old, new)
                        changed = True
    if not changed:
        raise AssertionError("未找到待改的 PPT 实际支出")
    presentation.save(path)


def remove_department_slide(path: Path) -> None:
    presentation = Presentation(path)
    slide_id = presentation.slides._sldIdLst[4]
    presentation.part.drop_rel(slide_id.rId)
    presentation.slides._sldIdLst.remove(slide_id)
    presentation.save(path)


def add_unsupported_assertion(path: Path) -> None:
    presentation = Presentation(path)
    slide = presentation.slides[-1]
    shape = slide.shapes.add_textbox(Inches(0.8), Inches(5.55), Inches(11.8), Inches(0.35))
    shape.text_frame.text = "Recommendation: will hire 500 employees next quarter."
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(12)
    presentation.save(path)


def add_full_slide_picture(path: Path, image_path: Path) -> None:
    presentation = Presentation(path)
    slide = presentation.slides[0]
    slide.shapes.add_picture(
        str(image_path),
        0,
        0,
        width=presentation.slide_width,
        height=presentation.slide_height,
    )
    presentation.save(path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--task-root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--result", type=Path)
    args = parser.parse_args()
    task_root = args.task_root.resolve()
    judge = load_judge(task_root)
    data_dir = task_root / "data" / "input_files"
    golden_dir = task_root / "solution"

    records: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="cowork-evo-budget-replay-") as temp_dir:
        temp = Path(temp_dir)

        fresh_output = temp / "fresh-output"
        run_solver(task_root, data_dir, fresh_output, temp / "fresh-workspace")
        fresh = judge.evaluate(fresh_output, data_dir)
        if not fresh["passed"] or fresh["score"] != 1.0:
            raise AssertionError(f"fresh replay 不是 1.0：{fresh}")
        records.append({"name": "fresh_replay", "expected": "accept", "score": fresh["score"]})

        mutated_data = temp / "mutated-data"
        shutil.copytree(data_dir, mutated_data)
        mutation = mutate_rd_salary(
            data_dir / "HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql",
            mutated_data / "HR_ANALYTICS__PUBLIC__EMPLOYEES.copy.sql",
        )
        mutated_output = temp / "mutated-output"
        run_solver(task_root, mutated_data, mutated_output, temp / "mutated-workspace")
        mutated = judge.evaluate(mutated_output, mutated_data)
        if not mutated["passed"] or mutated["score"] != 1.0:
            raise AssertionError(f"内容变异 replay 不是 1.0：{mutated}")
        mutated_rows, mutated_totals = judge.load_truth(mutated_data)
        rd = next(row for row in mutated_rows if row["department"] == "НИОКР")
        if rd["actual"] != Decimal("810147671.00") or rd["status"] != "over budget":
            raise AssertionError(f"变异后 НИОКР 真值异常：{rd}")
        if mutated_totals["most_under"] != "Кадры" or mutated_totals["most_over"] != "НИОКР":
            raise AssertionError(f"变异后排名没有联动：{mutated_totals}")
        records.append(
            {
                "name": "content_mutation_replay",
                "expected": "accept",
                "score": mutated["score"],
                "mutation": mutation,
                "changed_truth": {
                    "rd_actual": str(rd["actual"]),
                    "rd_status": rd["status"],
                    "most_under": mutated_totals["most_under"],
                    "most_over": mutated_totals["most_over"],
                    "total_actual": str(mutated_totals["total_actual"]),
                    "total_variance": str(mutated_totals["total_variance"]),
                },
            }
        )

        hardcoded = temp / "negative-hardcoded-processor"
        copy_outputs(golden_dir, hardcoded)
        hardcoded_result = expect_rejected(judge, hardcoded, mutated_data, "R003")
        records.append(
            {
                "name": "hardcoded_processor_ignores_changed_input",
                "expected": "reject",
                "score": hardcoded_result["score"],
                "hurdle_failures": hardcoded_result["hurdle_failures"],
            }
        )

        missing = temp / "negative-missing-file"
        missing.mkdir()
        shutil.copy2(golden_dir / "budget_vs_actuals.pptx", missing / "budget_vs_actuals.pptx")
        missing_result = expect_rejected(judge, missing, data_dir, "R001")
        records.append({"name": "missing_xlsx", "expected": "reject", "score": missing_result["score"]})

        wrong_excel = temp / "negative-wrong-excel-fact"
        copy_outputs(golden_dir, wrong_excel)
        change_excel_actual(wrong_excel / "budget_vs_actuals.xlsx")
        wrong_excel_result = expect_rejected(judge, wrong_excel, data_dir, "R003")
        records.append({"name": "wrong_excel_fact", "expected": "reject", "score": wrong_excel_result["score"]})

        broken_formula = temp / "negative-broken-formula"
        copy_outputs(golden_dir, broken_formula)
        break_excel_formula(broken_formula / "budget_vs_actuals.xlsx")
        broken_formula_result = expect_rejected(judge, broken_formula, data_dir, "R003")
        records.append({"name": "broken_variance_formula", "expected": "reject", "score": broken_formula_result["score"]})

        ppt_conflict = temp / "negative-cross-file-conflict"
        copy_outputs(golden_dir, ppt_conflict)
        change_ppt_actual(ppt_conflict / "budget_vs_actuals.pptx")
        ppt_conflict_result = expect_rejected(judge, ppt_conflict, data_dir, "R005")
        records.append({"name": "ppt_excel_conflict", "expected": "reject", "score": ppt_conflict_result["score"]})

        missing_slide = temp / "negative-missing-department-slide"
        copy_outputs(golden_dir, missing_slide)
        remove_department_slide(missing_slide / "budget_vs_actuals.pptx")
        missing_slide_result = expect_rejected(judge, missing_slide, data_dir, "R004")
        records.append({"name": "missing_department_slide", "expected": "reject", "score": missing_slide_result["score"]})

        unsupported = temp / "negative-unsupported-assertion"
        copy_outputs(golden_dir, unsupported)
        add_unsupported_assertion(unsupported / "budget_vs_actuals.pptx")
        unsupported_result = expect_rejected(judge, unsupported, data_dir, "R010")
        records.append({"name": "unsupported_assertion", "expected": "reject", "score": unsupported_result["score"]})

        flattened = temp / "negative-full-slide-picture"
        copy_outputs(golden_dir, flattened)
        add_full_slide_picture(
            flattened / "budget_vs_actuals.pptx",
            golden_dir / "qa" / "slides" / "slide-01.png",
        )
        flattened_result = expect_rejected(judge, flattened, data_dir, "R008")
        records.append({"name": "full_slide_picture", "expected": "reject", "score": flattened_result["score"]})

    result = {
        "task_id": judge.TASK_ID,
        "fresh_replay_score": 1.0,
        "content_mutation_replay_score": 1.0,
        "negative_cases": sum(1 for record in records if record["expected"] == "reject"),
        "all_expectations_met": True,
        "records": records,
    }
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(payload, encoding="utf-8")
    print(payload, end="")


if __name__ == "__main__":
    main()
