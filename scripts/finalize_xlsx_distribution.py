from __future__ import annotations

import csv
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TASKS = ROOT / "tasks" / "xlsx"
ARCHIVE = ROOT / "datasets" / "xlsx-archive"

ARCHIVE_TASKS = [
    ("sb2-visualization-129-mars-valuation", "intermediate", "结构与其他可视化/估值题重复，保留为可恢复候选"),
    ("sb2-template-14-09-concession-revenue", "intermediate", "单模板公式修复，与其他财务模板题重复，保留为可恢复候选"),
    ("officebench-2-13-section-calendar", "advanced", "日历单交付与其他 OfficeBench 日历题重复，保留为可恢复候选"),
    ("officebench-2-15-calendar-to-xlsx", "advanced", "日历到表格转换与其他跨应用日历题重复，保留为可恢复候选"),
    ("officebench-3-1-task-pdf-email", "advanced", "PDF/邮件单链路与其他跨应用通知题重复，保留为可恢复候选"),
    ("officebench-3-6-class-training-email", "advanced", "培训日历/邮件单链路与其他通知题重复，保留为可恢复候选"),
]

RETIER_EXPERT = {
    "cowork-bench-customer-satisfaction-analysis-report": "异构问卷与客服快照驱动两份 Excel、两份 Word，要求指标、行动计划和报告跨交付一致。",
    "cowork-bench-json-data-cleaning-pipeline": "两份 JSON 清洗、可回放脚本、两个 Excel 报告和进度状态形成跨文件数据管线。",
    "cowork-bench-moex-sector-rotation-dashboard": "市场价格、财报和研究门户三类输入汇入专业行业轮动分析看板，包含多源追溯与图表交付。",
    "officeval-082-checkbox-driven-tag-summary": "多工作表/多文件输入、交互式勾选控件、公式汇总和格式保真形成状态依赖工作流。",
}

def update_tier(task_id: str, rationale: str) -> None:
    root = TASKS / task_id
    toml = root / "task.toml"
    text = toml.read_text(encoding="utf-8").replace('difficulty_tier = "advanced"', 'difficulty_tier = "expert"')
    if "notes =" not in text:
        text = text.replace('publishability = "public"\n', 'publishability = "public"\nnotes = "author_hypothesis=L4; calibration_status=uncalibrated; review_status=pending_human_review"\n')
    else:
        text = text.replace("author_hypothesis=L3", "author_hypothesis=L4").replace("author_hypothesis=L2", "author_hypothesis=L4").replace("author_hypothesis=advanced", "author_hypothesis=L4")
    toml.write_text(text, encoding="utf-8")
    task_json = root / "metadata" / "task.json"
    data = json.loads(task_json.read_text(encoding="utf-8"))
    data["difficulty_tier"] = "expert"
    data["notes"] = f"author_hypothesis=L4; calibration_status=uncalibrated; review_status=pending_human_review; structural_rationale={rationale}"
    task_json.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    readme = root / "adaptation" / "README.md"
    with readme.open("a", encoding="utf-8") as f:
        f.write(f"\n## 难度档位调整记录（2026-08-31）\n\n本题从 L3/advanced 调整为 L4/expert 的作者结构假设，依据是：{rationale} 难度来自真实工作流中的异构输入、跨文件一致性、状态或专业分析链条，不是单次 Sonnet 5 得分。calibration_status=uncalibrated、review_status=pending_human_review 保持不变，最终需结合多模型结果、失败原因和人工复核。\n")

def main() -> None:
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    rows = []
    for task_id, old_tier, reason in ARCHIVE_TASKS:
        src = TASKS / task_id
        dst = ARCHIVE / task_id
        if src.exists():
            shutil.move(str(src), str(dst))
        rows.append({"task_id": task_id, "former_tier": old_tier, "archive_path": str(dst.relative_to(ROOT)), "reason": reason, "recoverable": "yes"})
    for task_id, rationale in RETIER_EXPERT.items():
        update_tier(task_id, rationale)
    manifest = ROOT / "datasets" / "xlsx-archive-manifest.csv"
    with manifest.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["task_id", "former_tier", "archive_path", "reason", "recoverable"])
        writer.writeheader()
        writer.writerows(rows)

if __name__ == "__main__":
    main()
