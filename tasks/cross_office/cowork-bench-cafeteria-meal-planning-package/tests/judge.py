#!/usr/bin/env python3
"""食堂菜单交付包的确定性 Judge；真值独立从输入重算。"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import date, timedelta
from html.parser import HTMLParser
from pathlib import Path

from docx import Document
from openpyxl import load_workbook


TASK_ID = "cowork-evo/cowork-bench-cafeteria-meal-planning-package"
RUBRIC_VERSION = "cowork-bench-cafeteria-meal-planning-package-v0.1"
JUDGE_VERSION = RUBRIC_VERSION
PASS_THRESHOLD = 0.8
PROCESSOR_NAME = "meal_cost_optimizer.py"
JSON_NAME = "meal_optimization_results.json"
XLSX_NAME = "Weekly_Meal_Budget.xlsx"
DOCX_NAME = "Cafeteria_Meal_Proposal.docx"
ICS_NAME = "meal_prep_calendar.ics"
OUTPUT_NAMES = [PROCESSOR_NAME, JSON_NAME, XLSX_NAME, DOCX_NAME, ICS_NAME]
STORES = ["FreshMart", "ValueGrocer", "BulkBarn"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
ALIASES = {
    "Говядина (вырезка)": "Говядина",
    "Филе рыбы (судак/треска)": "Филе рыбы",
    "Фарш мясной (свинина+говядина)": "Фарш мясной",
}
CRITERION_SPECS = {
    "R001": ("五个指定交付物完整可读", 5, "hurdle"),
    "R002": ("处理脚本在原始输入上回放正确", 7, "hurdle"),
    "R003": ("处理脚本在变异价格页上随输入变化", 7, "hurdle"),
    "R004": ("JSON 的 22 行价格与输入一致", 6, "hurdle"),
    "R005": ("JSON 的 5 道入选菜品与成本正确", 8, "hurdle"),
    "R006": ("JSON 菜单与预算汇总正确", 7, "hurdle"),
    "R007": ("Excel 三表结构和真值正确", 8, "hurdle"),
    "R008": ("Excel 公式可审计且重算正确", 7, "hurdle"),
    "R009": ("Word 提案内容完整正确", 6, "hurdle"),
    "R010": ("ICS 的 5 个备菜事件正确", 7, "hurdle"),
    "R011": ("四类业务交付物跨文件一致", 6, "hurdle"),
    "R012": ("Excel/Word 原生可编辑且格式专业", 4, "positive"),
    "R013": ("Excel/Word 可用 LibreOffice 重算与渲染", 3, "positive"),
    "R014": ("交付明确披露计价和营养数据限制", 4, "positive"),
    "R015": ("核心值与输入真值冲突", -15, "penalty"),
    "R016": ("处理脚本不可回放或硬编码", -14, "penalty"),
    "R017": ("交付缺失、损坏、截图化或跨文件冲突", -14, "penalty"),
    "R018": ("编造营养合规、未报价价格或重量换算", -18, "penalty"),
}
HURDLE_CAPS = {"R001": 0.0, "R002": 0.35, "R003": 0.45, "R004": 0.35, "R005": 0.35, "R006": 0.5, "R007": 0.4, "R008": 0.6, "R009": 0.55, "R010": 0.55, "R011": 0.65}


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows = []
        self.row = None
        self.cell = None

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "tr":
            self.row = []
        elif tag.lower() in {"td", "th"} and self.row is not None:
            self.cell = []

    def handle_data(self, data):
        if self.cell is not None:
            self.cell.append(data)

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag in {"td", "th"} and self.row is not None and self.cell is not None:
            self.row.append(" ".join("".join(self.cell).split()))
            self.cell = None
        elif tag == "tr" and self.row:
            self.rows.append(self.row)
            self.row = None


def criterion(identifier: str, passed: bool, evidence: str) -> dict:
    description, weight, kind = CRITERION_SPECS[identifier]
    return {"criterion_id": identifier, "score": 1.0 if passed else 0.0, "evidence": evidence, "method": "deterministic", "description": description, "weight": weight, "type": kind, "passed": passed}


def norm(value) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()


def formula_norm(value) -> str:
    return re.sub(r"[\s$']", "", str(value or "")).upper()


def close(left, right, tolerance=0.011) -> bool:
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def parse_number(value: str) -> float:
    match = re.search(r"-?\d+(?:[.,]\d+)?", value.replace("\xa0", " "))
    if not match:
        raise ValueError(value)
    return float(match.group(0).replace(",", "."))


def derive_expected(prices_path: Path, recipes_path: Path) -> dict:
    parser = TableParser()
    parser.feed(prices_path.read_text(encoding="utf-8"))
    if not parser.rows or parser.rows[0][2:5] != STORES:
        raise ValueError("价格表头不符合预期")
    prices = []
    for raw in parser.rows[1:]:
        if len(raw) < 5:
            continue
        values = {store: parse_number(raw[index + 2]) for index, store in enumerate(STORES)}
        minimum = min(values.values())
        prices.append({
            "Item": raw[0], "Unit": raw[1],
            **{f"{store}_Price": values[store] for store in STORES},
            "Cheapest_Store": next(store for store in STORES if values[store] == minimum),
            "Cheapest_Price": minimum,
        })
    prices.sort(key=lambda row: row["Item"])
    if len(prices) != 22:
        raise ValueError(f"价格行数={len(prices)}")
    by_item = {row["Item"]: row for row in prices}
    recipes = json.loads(recipes_path.read_text(encoding="utf-8"))
    if len(recipes) != 50:
        raise ValueError(f"配方行数={len(recipes)}")
    candidates = []
    for recipe in recipes:
        if recipe.get("category") != "горячее":
            continue
        matched, unpriced, seen = [], [], set()
        for ingredient in recipe.get("ingredients", []):
            source_name = str(ingredient.get("name") or "").strip()
            price_item = ALIASES.get(source_name, source_name)
            if price_item in by_item and price_item not in seen:
                row = by_item[price_item]
                matched.append({"Recipe_Ingredient": source_name, "Price_Item": price_item, "Cheapest_Store": row["Cheapest_Store"], "Quoted_Price": row["Cheapest_Price"]})
                seen.add(price_item)
            elif price_item not in by_item:
                unpriced.append(source_name)
        if len(matched) < 3:
            continue
        total = round(sum(row["Quoted_Price"] for row in matched), 2)
        candidates.append({
            "Recipe_ID": recipe["id"], "Recipe": recipe["name"],
            "Matched_Ingredient_Count": len(matched), "Unpriced_Ingredient_Count": len(unpriced),
            "Matched_Ingredients": matched, "Unpriced_Ingredients": unpriced,
            "Total_Cost": total,
            "Cheapest_Store_Mix": [store for store in STORES if any(row["Cheapest_Store"] == store for row in matched)],
            "Cost_Per_Serving": round(total / 4.0, 2),
        })
    selected = sorted(candidates, key=lambda row: (row["Total_Cost"], row["Recipe_ID"]))[:5]
    plan = []
    start = date(2026, 3, 16)
    reverse = list(reversed(selected))
    for index, day in enumerate(DAYS):
        current = (start + timedelta(days=index)).isoformat()
        for meal, recipe in (("Lunch", selected[index]), ("Dinner", reverse[index])):
            plan.append({"Date": current, "Day": day, "Meal": meal, "Recipe_ID": recipe["Recipe_ID"], "Recipe": recipe["Recipe"], "Estimated_Cost": recipe["Total_Cost"]})
    selected_by_id = {row["Recipe_ID"]: row for row in selected}
    single_store = {}
    for store in STORES:
        total = 0.0
        for meal in plan:
            for ingredient in selected_by_id[meal["Recipe_ID"]]["Matched_Ingredients"]:
                total += by_item[ingredient["Price_Item"]][f"{store}_Price"]
        single_store[store] = round(total, 2)
    weekly = round(sum(row["Estimated_Cost"] for row in plan), 2)
    best_store = min(STORES, key=lambda store: (single_store[store], STORES.index(store)))
    best_total = single_store[best_store]
    savings = round(best_total - weekly, 2)
    budget = {
        "Weekly_Total": weekly, "Average_Daily": round(weekly / 5.0, 2), "Weekly_Budget": 12000.0,
        "Budget_Headroom": round(12000.0 - weekly, 2), "Best_Single_Store": best_store,
        "Best_Single_Store_Total": best_total, "Savings_Vs_Best_Single_Store": savings,
        "Savings_Pct": round(savings / best_total * 100.0, 2), "Single_Store_Totals": single_store,
    }
    return {"price_rows": prices, "selected_recipes": selected, "weekly_meal_plan": plan, "budget_summary": budget}


def compare_price_rows(actual, expected) -> list[str]:
    if not isinstance(actual, list) or len(actual) != len(expected):
        return [f"价格行数错误: {len(actual) if isinstance(actual, list) else type(actual).__name__}"]
    errors = []
    fields = ["Item", "Unit", "FreshMart_Price", "ValueGrocer_Price", "BulkBarn_Price", "Cheapest_Store", "Cheapest_Price"]
    for index, (got, want) in enumerate(zip(actual, expected), start=1):
        if not isinstance(got, dict):
            errors.append(f"价格第{index}行非对象"); continue
        for field in fields:
            valid = close(got.get(field), want[field]) if field.endswith("_Price") else str(got.get(field, "")).strip() == str(want[field])
            if not valid:
                errors.append(f"价格第{index}行 {field}: {got.get(field)!r} != {want[field]!r}")
    return errors


def compare_selected(actual, expected) -> list[str]:
    if not isinstance(actual, list) or len(actual) != 5:
        return [f"入选菜品数错误: {len(actual) if isinstance(actual, list) else type(actual).__name__}"]
    errors = []
    scalar_fields = ["Recipe_ID", "Recipe", "Matched_Ingredient_Count", "Unpriced_Ingredient_Count", "Total_Cost", "Cost_Per_Serving"]
    for index, (got, want) in enumerate(zip(actual, expected), start=1):
        if not isinstance(got, dict):
            errors.append(f"菜品第{index}行非对象"); continue
        for field in scalar_fields:
            valid = close(got.get(field), want[field]) if field in {"Total_Cost", "Cost_Per_Serving"} else got.get(field) == want[field]
            if not valid:
                errors.append(f"菜品第{index}行 {field}: {got.get(field)!r} != {want[field]!r}")
        if got.get("Matched_Ingredients") != want["Matched_Ingredients"]:
            errors.append(f"菜品第{index}行 Matched_Ingredients 不一致")
        if got.get("Unpriced_Ingredients") != want["Unpriced_Ingredients"]:
            errors.append(f"菜品第{index}行 Unpriced_Ingredients 不一致")
        if got.get("Cheapest_Store_Mix") != want["Cheapest_Store_Mix"]:
            errors.append(f"菜品第{index}行 Cheapest_Store_Mix 不一致")
    return errors


def compare_plan(actual, expected) -> list[str]:
    if not isinstance(actual, list) or len(actual) != 10:
        return [f"菜单行数错误: {len(actual) if isinstance(actual, list) else type(actual).__name__}"]
    errors = []
    for index, (got, want) in enumerate(zip(actual, expected), start=1):
        for field in ("Date", "Day", "Meal", "Recipe_ID", "Recipe"):
            if str(got.get(field, "")).strip() != str(want[field]):
                errors.append(f"菜单第{index}行 {field} 错误")
        if not close(got.get("Estimated_Cost"), want["Estimated_Cost"]):
            errors.append(f"菜单第{index}行 Estimated_Cost 错误")
    return errors


def compare_budget(actual, expected) -> list[str]:
    if not isinstance(actual, dict):
        return ["预算汇总不是对象"]
    errors = []
    for field in ("Weekly_Total", "Average_Daily", "Weekly_Budget", "Budget_Headroom", "Best_Single_Store_Total", "Savings_Vs_Best_Single_Store", "Savings_Pct"):
        if not close(actual.get(field), expected[field]):
            errors.append(f"{field}: {actual.get(field)!r} != {expected[field]!r}")
    if actual.get("Best_Single_Store") != expected["Best_Single_Store"]:
        errors.append("Best_Single_Store 错误")
    totals = actual.get("Single_Store_Totals")
    if not isinstance(totals, dict) or any(not close(totals.get(store), expected["Single_Store_Totals"][store]) for store in STORES):
        errors.append("Single_Store_Totals 错误")
    return errors


def run_processor(script: Path, prices: Path, recipes: Path, output: Path) -> tuple[dict | None, str]:
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        result = subprocess.run(
            [sys.executable, str(script.resolve()), "--prices", str(prices.resolve()), "--recipes", str(recipes.resolve()), "--output", str(output.resolve())],
            cwd=str(script.parent.resolve()), capture_output=True, text=True, timeout=90, check=False,
        )
    except Exception as error:
        return None, f"执行异常: {error}"
    if result.returncode != 0 or not output.is_file():
        return None, f"exit={result.returncode}; stderr={(result.stderr or '')[-600:]}"
    try:
        return json.loads(output.read_text(encoding="utf-8")), f"exit=0，输出={output.name}"
    except Exception as error:
        return None, f"输出 JSON 无法读取: {error}"


def mutate_prices(source: Path, destination: Path) -> None:
    original = source.read_text(encoding="utf-8")
    needle = '<tr><td>Крупа гречневая</td><td>за кг</td><td class="price">109 ₽</td>'
    replacement = '<tr><td>Крупа гречневая</td><td>за кг</td><td class="price">500 ₽</td>'
    if original.count(needle) != 1:
        raise ValueError("无法唯一定位变异报价")
    destination.write_text(original.replace(needle, replacement), encoding="utf-8")


def extract_rows(sheet, headers: list[str]) -> tuple[list[dict], list[str]]:
    actual_headers = [str(cell.value or "").strip() for cell in sheet[1]]
    missing = [header for header in headers if header not in actual_headers]
    if missing:
        return [], missing
    positions = {header: actual_headers.index(header) for header in headers}
    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        rows.append({header: row[positions[header]] for header in headers})
    return rows, []


def libreoffice_prepare(xlsx: Path, docx: Path, root: Path):
    libreoffice, pdftotext, pdfinfo = shutil.which("libreoffice"), shutil.which("pdftotext"), shutil.which("pdfinfo")
    if not all((libreoffice, pdftotext, pdfinfo)):
        return None, {}, {}, "缺少 libreoffice/pdftotext/pdfinfo"
    source_dir, recalc_dir, pdf_dir = root / "source", root / "recalc", root / "pdf"
    source_dir.mkdir(); recalc_dir.mkdir(); pdf_dir.mkdir()
    xlsx_copy, docx_copy = source_dir / XLSX_NAME, source_dir / DOCX_NAME
    shutil.copy2(xlsx, xlsx_copy); shutil.copy2(docx, docx_copy)
    profile_uri = (root / "lo-profile").resolve().as_uri()
    recalc = subprocess.run([libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "xlsx", "--outdir", str(recalc_dir), str(xlsx_copy)], capture_output=True, text=True, timeout=300, check=False)
    recalculated = recalc_dir / XLSX_NAME
    if recalc.returncode != 0 or not recalculated.is_file():
        return None, {}, {}, (recalc.stderr or recalc.stdout or "Excel 重算失败")[-800:]
    texts, pages = {}, {}
    for source in (recalculated, docx_copy):
        converted = subprocess.run([libreoffice, "--headless", f"-env:UserInstallation={profile_uri}", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(source)], capture_output=True, text=True, timeout=300, check=False)
        pdf_path = pdf_dir / f"{source.stem}.pdf"
        if converted.returncode != 0 or not pdf_path.is_file():
            return recalculated, texts, pages, (converted.stderr or converted.stdout or f"{source.name} 渲染失败")[-800:]
        extracted = subprocess.run([pdftotext, "-raw", str(pdf_path), "-"], capture_output=True, text=True, timeout=120, check=False)
        info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=120, check=False)
        if extracted.returncode != 0 or info.returncode != 0:
            return recalculated, texts, pages, f"{source.name} PDF 检查失败"
        match = re.search(r"^Pages:\s+(\d+)", info.stdout, re.MULTILINE)
        texts[source.name], pages[source.name] = extracted.stdout, int(match.group(1)) if match else 0
    return recalculated, texts, pages, "Excel 重算与 Excel/Word PDF 渲染成功"


def workbook_formatting(workbook) -> dict:
    result = {}
    for sheet in workbook.worksheets:
        result[sheet.title] = {
            "tables": len(sheet.tables), "freeze": str(sheet.freeze_panes or ""),
            "header_bold": all(bool(cell.font.bold) for cell in sheet[1] if cell.value is not None),
            "header_fill": all(cell.fill.fill_type == "solid" for cell in sheet[1] if cell.value is not None),
            "wide_columns": sum(1 for dimension in sheet.column_dimensions.values() if (dimension.width or 0) >= 12) >= min(3, sheet.max_column),
        }
    return result


def document_snapshot(path: Path):
    document = Document(path)
    paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    table_rows = [[cell.text.strip() for cell in row.cells] for table in document.tables for row in table.rows]
    headings = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.style and paragraph.style.name.startswith("Heading")]
    with zipfile.ZipFile(path) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
    return document, "\n".join(paragraphs + [" | ".join(row) for row in table_rows]), headings, len(table_rows), media


def unfold_ics(text: str) -> list[str]:
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    unfolded = []
    for line in lines:
        if line.startswith((" ", "\t")) and unfolded:
            unfolded[-1] += line[1:]
        elif line:
            unfolded.append(line)
    return unfolded


def unescape_ics(value: str) -> str:
    return value.replace("\\n", "\n").replace("\\,", ",").replace("\\;", ";").replace("\\\\", "\\")


def parse_ics(path: Path) -> list[dict]:
    events, current = [], None
    for line in unfold_ics(path.read_text(encoding="utf-8")):
        if line == "BEGIN:VEVENT":
            current = {}
        elif line == "END:VEVENT" and current is not None:
            events.append(current); current = None
        elif current is not None and ":" in line:
            key, value = line.split(":", 1)
            current[key] = unescape_ics(value)
    return events


def unsupported_claims(text: str) -> list[str]:
    patterns = {
        "protein_verified": r"(?:содержит|обеспечивает|даёт|дает)\s*(?:не\s+менее\s+|минимум\s+)?20\s*(?:г|g)\b",
        "protein_compliant": r"(?:protein|белок).{0,50}(?:compliant|requirement met|требование\s+выполнено|норма\s+соблюдена)",
    }
    return [name for name, pattern in patterns.items() if re.search(pattern, text, re.IGNORECASE | re.DOTALL)]


def evaluate(artifact_dir: Path) -> dict:
    paths = {name: artifact_dir / name for name in OUTPUT_NAMES}
    results = []
    if not all(path.is_file() for path in paths.values()):
        results.append(criterion("R001", False, "文件存在性=" + json.dumps({name: path.is_file() for name, path in paths.items()}, ensure_ascii=False)))
        return build_result(results, list(paths.values()))
    try:
        delivered = json.loads(paths[JSON_NAME].read_text(encoding="utf-8"))
        script_text = paths[PROCESSOR_NAME].read_text(encoding="utf-8")
        ics_text = paths[ICS_NAME].read_text(encoding="utf-8")
        if len(script_text.strip()) < 500 or not zipfile.is_zipfile(paths[XLSX_NAME]) or not zipfile.is_zipfile(paths[DOCX_NAME]):
            raise ValueError("脚本过短或 Office 文件非有效 OOXML")
        if "BEGIN:VCALENDAR" not in ics_text or "END:VCALENDAR" not in ics_text:
            raise ValueError("ICS 缺少 VCALENDAR 边界")
    except Exception as error:
        results.append(criterion("R001", False, f"交付物无法读取: {error}"))
        return build_result(results, list(paths.values()))
    results.append(criterion("R001", True, f"五文件可读；脚本字符={len(script_text)}"))

    input_dir = Path(__file__).resolve().parents[1] / "data" / "input_files"
    expected = derive_expected(input_dir / "grocery_prices.html", input_dir / "all_recipes.json")
    with tempfile.TemporaryDirectory(prefix="cowork-evo-meal-script-") as temp_name:
        temp = Path(temp_name)
        canonical, canonical_evidence = run_processor(paths[PROCESSOR_NAME], input_dir / "grocery_prices.html", input_dir / "all_recipes.json", temp / "canonical.json")
        canonical_errors = [] if canonical is not None else [canonical_evidence]
        if canonical is not None:
            canonical_errors = compare_price_rows(canonical.get("price_rows"), expected["price_rows"]) + compare_selected(canonical.get("selected_recipes"), expected["selected_recipes"]) + compare_plan(canonical.get("weekly_meal_plan"), expected["weekly_meal_plan"]) + compare_budget(canonical.get("budget_summary"), expected["budget_summary"])
        canonical_ok = not canonical_errors
        results.append(criterion("R002", canonical_ok, f"{canonical_evidence}，错误={canonical_errors[:8]}"))

        mutated_path = temp / "mutated-prices.html"
        mutate_prices(input_dir / "grocery_prices.html", mutated_path)
        mutated_expected = derive_expected(mutated_path, input_dir / "all_recipes.json")
        mutated, mutated_evidence = run_processor(paths[PROCESSOR_NAME], mutated_path, input_dir / "all_recipes.json", temp / "mutated.json")
        mutated_errors = [] if mutated is not None else [mutated_evidence]
        if mutated is not None:
            mutated_errors = compare_price_rows(mutated.get("price_rows"), mutated_expected["price_rows"]) + compare_selected(mutated.get("selected_recipes"), mutated_expected["selected_recipes"]) + compare_plan(mutated.get("weekly_meal_plan"), mutated_expected["weekly_meal_plan"]) + compare_budget(mutated.get("budget_summary"), mutated_expected["budget_summary"])
        changed = bool(mutated and canonical and mutated.get("budget_summary") != canonical.get("budget_summary"))
        mutated_ok = not mutated_errors and changed
        results.append(criterion("R003", mutated_ok, f"{mutated_evidence}，错误={mutated_errors[:8]}，预算已改变={changed}，变异周总额={mutated_expected['budget_summary']['Weekly_Total']}"))

    price_errors = compare_price_rows(delivered.get("price_rows"), expected["price_rows"])
    recipe_errors = compare_selected(delivered.get("selected_recipes"), expected["selected_recipes"])
    plan_errors = compare_plan(delivered.get("weekly_meal_plan"), expected["weekly_meal_plan"])
    budget_errors = compare_budget(delivered.get("budget_summary"), expected["budget_summary"])
    results.append(criterion("R004", not price_errors, f"价格错误={price_errors[:10]}"))
    results.append(criterion("R005", not recipe_errors, f"菜品错误={recipe_errors[:10]}"))
    daily = {}
    if isinstance(delivered.get("weekly_meal_plan"), list):
        for row in delivered["weekly_meal_plan"]:
            if isinstance(row, dict):
                daily[row.get("Date")] = daily.get(row.get("Date"), 0) + float(row.get("Estimated_Cost") or 0)
    daily_ok = len(daily) == 5 and all(value < 2000.0 for value in daily.values())
    results.append(criterion("R006", not plan_errors and not budget_errors and daily_ok, f"菜单错误={plan_errors[:8]}，预算错误={budget_errors[:8]}，日成本={daily}"))

    with tempfile.TemporaryDirectory(prefix="cowork-evo-meal-office-") as temp_name:
        recalculated, rendered_texts, rendered_pages, render_evidence = libreoffice_prepare(paths[XLSX_NAME], paths[DOCX_NAME], Path(temp_name))
        try:
            workbook_formula = load_workbook(paths[XLSX_NAME], data_only=False)
            workbook_values = load_workbook(recalculated or paths[XLSX_NAME], data_only=True)
            formatting = workbook_formatting(workbook_formula)
        except Exception as error:
            workbook_formula = workbook_values = None
            formatting = {"error": str(error)}

        expected_sheets = ["Grocery_Prices", "Recipe_Cost_Analysis", "Weekly_Meal_Plan"]
        sheet_ok = workbook_formula is not None and workbook_formula.sheetnames == expected_sheets and workbook_values.sheetnames == expected_sheets
        excel_price_errors, excel_recipe_errors, excel_plan_errors = ["Excel 无法读取"], ["Excel 无法读取"], ["Excel 无法读取"]
        price_rows = recipe_rows = meal_rows = []
        if sheet_ok:
            price_headers = ["Item", "Unit", "FreshMart_Price", "ValueGrocer_Price", "BulkBarn_Price", "Cheapest_Store", "Cheapest_Price"]
            recipe_headers = ["Recipe_ID", "Recipe", "Matched_Ingredient_Count", "Unpriced_Ingredient_Count", "Total_Cost", "Cheapest_Store_Mix", "Cost_Per_Serving", "Unpriced_Ingredients"]
            plan_headers = ["Date", "Day", "Meal", "Recipe_ID", "Recipe", "Estimated_Cost"]
            price_rows, missing_price = extract_rows(workbook_values["Grocery_Prices"], price_headers)
            recipe_rows, missing_recipe = extract_rows(workbook_values["Recipe_Cost_Analysis"], recipe_headers)
            meal_rows, missing_plan = extract_rows(workbook_values["Weekly_Meal_Plan"], plan_headers)
            excel_price_errors = missing_price or compare_price_rows(price_rows, expected["price_rows"])
            projected_recipes = []
            if not missing_recipe:
                for row, truth in zip(recipe_rows, expected["selected_recipes"]):
                    projected_recipes.append({
                        **row,
                        "Matched_Ingredients": truth["Matched_Ingredients"],
                        "Unpriced_Ingredients": [part.strip() for part in str(row["Unpriced_Ingredients"] or "").split(";") if part.strip()],
                        "Cheapest_Store_Mix": [part for part in str(row["Cheapest_Store_Mix"] or "").split("/") if part],
                    })
            excel_recipe_errors = missing_recipe or compare_selected(projected_recipes, expected["selected_recipes"])
            excel_plan_errors = missing_plan or compare_plan(meal_rows, expected["weekly_meal_plan"])
        excel_truth_ok = sheet_ok and not excel_price_errors and not excel_recipe_errors and not excel_plan_errors
        results.append(criterion("R007", excel_truth_ok, f"工作表={workbook_formula.sheetnames if workbook_formula else None}，价格错={excel_price_errors[:5]}，菜品错={excel_recipe_errors[:5]}，菜单错={excel_plan_errors[:5]}"))

        formula_checks, formulas = [], {}
        if workbook_formula is not None and sheet_ok:
            for row in range(2, 24):
                f, g = formula_norm(workbook_formula["Grocery_Prices"][f"F{row}"].value), formula_norm(workbook_formula["Grocery_Prices"][f"G{row}"].value)
                formulas[f"Grocery_Prices!F{row}"] = f; formulas[f"Grocery_Prices!G{row}"] = g
                formula_checks.extend([f.startswith("=") and "MATCH(MIN(" in f and "INDEX(" in f, g.startswith("=MIN(")])
            for row in range(2, 7):
                g = formula_norm(workbook_formula["Recipe_Cost_Analysis"][f"G{row}"].value)
                formulas[f"Recipe_Cost_Analysis!G{row}"] = g
                formula_checks.append(g.startswith("=ROUND(") and f"E{row}/4" in g)
            for row in range(2, 12):
                f = formula_norm(workbook_formula["Weekly_Meal_Plan"][f"F{row}"].value)
                formulas[f"Weekly_Meal_Plan!F{row}"] = f
                formula_checks.append(f.startswith("=") and "RECIPE_COST_ANALYSIS!" in f and ("VLOOKUP(" in f or "XLOOKUP(" in f or "INDEX(" in f))
        formula_ok = len(formula_checks) == 59 and all(formula_checks) and excel_truth_ok and recalculated is not None
        results.append(criterion("R008", formula_ok, f"公式检查={sum(formula_checks) if formula_checks else 0}/{len(formula_checks)}，LibreOffice重算={recalculated is not None}，样例={dict(list(formulas.items())[:5])}"))

        try:
            document, doc_text, headings, table_row_count, media = document_snapshot(paths[DOCX_NAME])
            doc_norm = norm(doc_text)
            required_headings = {"budget overview", "menu highlights", "procurement strategy", "limitations"}
            heading_ok = required_headings.issubset({norm(item) for item in headings})
            recipe_missing_doc = [item["Recipe"] for item in expected["selected_recipes"] if norm(item["Recipe"]) not in doc_norm]
            number_tokens = ["7,632", "1,526.40", "4,368", "7,946", "314", "3.95"]
            normalized_digits = re.sub(r"\s", "", doc_text)
            number_missing = [token for token in number_tokens if token not in normalized_digits]
            doc_content_ok = "weekly cafeteria meal proposal" in doc_norm and heading_ok and not recipe_missing_doc and not number_missing and all(norm(store) in doc_norm for store in STORES) and "quoted-basket" in doc_norm and table_row_count >= 6
        except Exception as error:
            doc_text, headings, table_row_count, media = "", [], 0, []
            recipe_missing_doc, number_missing = [str(error)], [str(error)]
            doc_content_ok = False
        results.append(criterion("R009", doc_content_ok, f"标题={headings}，缺菜={recipe_missing_doc}，缺数字={number_missing}，表行={table_row_count}"))

        try:
            events = parse_ics(paths[ICS_NAME])
            lunches = [row for row in expected["weekly_meal_plan"] if row["Meal"] == "Lunch"]
            event_errors = []
            if len(events) != 5:
                event_errors.append(f"事件数={len(events)}")
            for index, (event, lunch) in enumerate(zip(events, lunches), start=1):
                stamp = lunch["Date"].replace("-", "")
                if event.get("DTSTART;TZID=Europe/Moscow") != f"{stamp}T080000": event_errors.append(f"事件{index}开始时间错")
                if event.get("DTEND;TZID=Europe/Moscow") != f"{stamp}T093000": event_errors.append(f"事件{index}结束时间错")
                if event.get("SUMMARY") != "Meal Prep: " + lunch["Recipe"]: event_errors.append(f"事件{index}标题错")
                description = event.get("DESCRIPTION", "")
                if not any(store in description for store in STORES): event_errors.append(f"事件{index}无商店")
                if "Без котировки" not in description: event_errors.append(f"事件{index}未披露未报价原料")
            ics_ok = not event_errors and len(events) == 5
        except Exception as error:
            events, event_errors, ics_ok = [], [str(error)], False
        results.append(criterion("R010", ics_ok, f"事件数={len(events)}，错误={event_errors}"))

        cross_ok = not (price_errors or recipe_errors or plan_errors or budget_errors or excel_price_errors or excel_recipe_errors or excel_plan_errors) and doc_content_ok and ics_ok
        results.append(criterion("R011", cross_ok, f"JSON真值={not (price_errors or recipe_errors or plan_errors or budget_errors)}，Excel真值={excel_truth_ok}，Word={doc_content_ok}，ICS={ics_ok}"))

        formatting_ok = isinstance(formatting, dict) and len(formatting) == 3 and all(item.get("tables", 0) >= 1 and item.get("freeze") == "A2" and item.get("header_bold") and item.get("header_fill") and item.get("wide_columns") for item in formatting.values())
        native_ok = formatting_ok and not media and table_row_count >= 6
        results.append(criterion("R012", native_ok, f"Excel样式={json.dumps(formatting, ensure_ascii=False)}，Word媒体={media}，表行={table_row_count}"))

        xlsx_render_norm = norm(rendered_texts.get(XLSX_NAME, ""))
        docx_render_norm = norm(rendered_texts.get(DOCX_NAME, ""))
        xlsx_render_ok = rendered_pages.get(XLSX_NAME, 0) == 3 and all(token in xlsx_render_norm for token in ("item", "cheapest_store", "recipe_id", "cost_per_serving", "estimated_cost", "гречка с тушёнкой"))
        docx_render_ok = rendered_pages.get(DOCX_NAME, 0) >= 2 and all(token in docx_render_norm for token in ("weekly cafeteria meal proposal", "budget overview", "procurement strategy", "limitations"))
        render_ok = recalculated is not None and xlsx_render_ok and docx_render_ok
        results.append(criterion("R013", render_ok, f"{render_evidence}，页数={rendered_pages}，Excel文本={xlsx_render_ok}，Word文本={docx_render_ok}"))

    limitations = delivered.get("limitations")
    limitation_text = norm(" ".join(str(item) for item in limitations)) if isinstance(limitations, list) else ""
    doc_lower = norm(doc_text)
    limitations_ok = "quoted-basket" in limitation_text and "cannot be verified" in limitation_text and "невозможно проверить" in doc_lower and "не импутировал" in doc_lower
    results.append(criterion("R014", limitations_ok, f"JSON限制={limitation_text[:260]}，Word限制={limitations_ok}"))

    unsupported = unsupported_claims(json.dumps(delivered, ensure_ascii=False) + "\n" + doc_text)
    source_error = bool(price_errors or recipe_errors or plan_errors or budget_errors or excel_price_errors or excel_recipe_errors or excel_plan_errors)
    script_error = not (canonical_ok and mutated_ok)
    broken_or_conflict = not (native_ok and render_ok and cross_ok)
    results.append(criterion("R015", source_error, f"JSON价格={not price_errors}，菜品={not recipe_errors}，菜单={not plan_errors}，预算={not budget_errors}，Excel={excel_truth_ok}"))
    results.append(criterion("R016", script_error, f"原始回放={canonical_ok}，变异回放={mutated_ok}"))
    results.append(criterion("R017", broken_or_conflict, f"原生={native_ok}，渲染={render_ok}，跨文件={cross_ok}"))
    results.append(criterion("R018", bool(unsupported), f"无来源断言={unsupported}"))
    return build_result(results, list(paths.values()))


def build_result(results: list[dict], artifact_paths: list[Path]) -> dict:
    present = {item["criterion_id"] for item in results}
    for identifier in CRITERION_SPECS:
        if identifier not in present:
            results.append(criterion(identifier, False, "因交付不存在、损坏或门槛失败，本判据未执行"))
    results.sort(key=lambda item: item["criterion_id"])
    positive_total = sum(item["weight"] for item in results if item["weight"] > 0)
    earned = sum(item["weight"] for item in results if item["weight"] > 0 and item["passed"])
    penalties = sum(item["weight"] for item in results if item["weight"] < 0 and item["passed"])
    reward = max(0.0, min(1.0, (earned + penalties) / positive_total))
    hurdle_failures = [item["criterion_id"] for item in results if item["type"] == "hurdle" and not item["passed"]]
    if hurdle_failures:
        reward = min(reward, min(HURDLE_CAPS[identifier] for identifier in hurdle_failures))
    digests = [hashlib.sha256(path.read_bytes()).hexdigest() for path in artifact_paths if path.is_file()] or [hashlib.sha256(b"").hexdigest()]
    criterion_results = []
    for item in results:
        score = (0.0 if item["passed"] else 1.0) if item["type"] == "penalty" else item["score"]
        criterion_results.append({
            "criterion_id": item["criterion_id"], "score": score, "evidence": item["evidence"], "method": item["method"],
            "raw": {"description": item["description"], "weight": item["weight"], "type": item["type"], "passed": item["passed"]},
        })
    return {"task_id": TASK_ID, "rubric_version": RUBRIC_VERSION, "judge_version": JUDGE_VERSION, "artifact_sha256": digests, "criterion_results": criterion_results, "reward": round(reward, 6), "passed": reward >= PASS_THRESHOLD, "hurdle_failures": hurdle_failures}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifact-dir", type=Path, required=True)
    parser.add_argument("--result", type=Path)
    parser.add_argument("--reward-file", type=Path)
    args = parser.parse_args()
    result = evaluate(args.artifact_dir)
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    print(rendered)
    if args.result:
        args.result.parent.mkdir(parents=True, exist_ok=True)
        args.result.write_text(rendered + "\n", encoding="utf-8")
    if args.reward_file:
        args.reward_file.parent.mkdir(parents=True, exist_ok=True)
        args.reward_file.write_text(f"{result['reward']:.6f}\n", encoding="utf-8")
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
