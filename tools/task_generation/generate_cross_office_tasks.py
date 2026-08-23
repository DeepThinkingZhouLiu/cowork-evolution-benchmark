"""生成 8 个跨 Office 任务家族的公开端到端样例。"""

from __future__ import annotations

from pathlib import Path

from common import (
    REFERENCE_URLS,
    create_docx,
    create_pdf,
    create_pptx,
    create_xlsx,
    write_csv,
    write_json,
    write_task_package,
    write_text,
)


ROOT = Path(__file__).resolve().parents[2]
TASKS_ROOT = ROOT / "tasks/cross_office"


def metadata(
    family: str,
    *,
    operations: list[str],
    modalities: list[str],
    scope: str,
    constraints: list[str],
    capabilities: list[str],
    sector: str,
    occupation: str,
    difficulty: str = "advanced",
    minutes: int = 75,
    outputs: list[str],
) -> dict:
    return {
        "domain": "cross_office",
        "task_family": family,
        "operation_types": operations,
        "input_modalities": modalities,
        "dependency_scope": scope,
        "constraint_profiles": constraints,
        "deliverable_types": outputs,
        "capability_ids": capabilities,
        "sector": sector,
        "occupation": occupation,
        "difficulty_tier": difficulty,
        "estimated_minutes": minutes,
        "multimodal": "image" in modalities,
        "additional_judges": ["semantic", "visual"],
    }


def instruction(title: str, context: str, inputs: list[str], outputs: list[str], requirements: list[str]) -> str:
    return "\n".join(
        [
            f"# {title}",
            "",
            context,
            "",
            "## 输入文件",
            "",
            *[f"- `{name}`" for name in inputs],
            "",
            "## 交付要求",
            "",
            *[f"- {item}" for item in requirements],
            "",
            "## 输出位置",
            "",
            *[f"- `/app/output/{name}`" for name in outputs],
            "",
            "不得联网。多个交付物中的同一数字、日期、责任人和结论必须完全一致；不得把输入材料中没有出现的内容写成事实。",
        ]
    )


def build_extract_analyze_present() -> None:
    task_id = "cross-001-extract-analyze-present"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    write_csv(
        input_dir / "门店经营明细.csv",
        [
            ["门店", "区域", "收入_元", "目标_元", "退货率", "满意度"],
            ["华东一店", "华东", 1250000, 1200000, "3.1%", 91],
            ["华东二店", "华东", 980000, 1000000, "4.2%", 86],
            ["华北一店", "华北", 870000, 900000, "5.4%", 83],
            ["华北二店", "华北", 640000, 850000, "7.2%", 75],
        ],
    )
    create_pdf(
        input_dir / "management_questions.pdf",
        [
            "1. Which store is furthest below target?",
            "2. Where is the return-rate risk concentrated?",
            "3. Recommend two actions with owners and due dates.",
            "4. Keep workbook and presentation facts consistent.",
        ],
        title="MANAGEMENT REVIEW QUESTIONS",
    )
    outputs = ["门店经营分析.xlsx", "门店经营复盘.pptx"]
    solve_spec = {
        "required_inputs": ["门店经营明细.csv", "management_questions.pdf"],
        "workbooks": [
            {
                "filename": outputs[0],
                "sheets": [
                    {
                        "name": "门店明细",
                        "rows": [
                            ["门店", "区域", "收入_元", "目标_元", "达成率", "退货率", "满意度"],
                            ["华东一店", "华东", 1250000, 1200000, None, "3.1%", 91],
                            ["华东二店", "华东", 980000, 1000000, None, "4.2%", 86],
                            ["华北一店", "华北", 870000, 900000, None, "5.4%", 83],
                            ["华北二店", "华北", 640000, 850000, None, "7.2%", 75],
                        ],
                        "formulas": {"E2": "=C2/D2", "E3": "=C3/D3", "E4": "=C4/D4", "E5": "=C5/D5"},
                        "chart": {"type": "bar", "title": "门店收入", "data_col": 3, "category_col": 1, "anchor": "I2"},
                    },
                    {
                        "name": "区域汇总",
                        "rows": [["区域", "收入_元", "目标_元", "达成率"], ["华东", 2230000, 2200000, None], ["华北", 1510000, 1750000, None], ["合计", 3740000, 3950000, None]],
                        "formulas": {"D2": "=B2/C2", "D3": "=B3/C3", "D4": "=B4/C4"},
                        "chart": {"type": "bar", "title": "区域收入对比", "data_col": 2, "category_col": 1, "anchor": "F2"},
                    },
                    {
                        "name": "行动清单",
                        "rows": [["优先级", "问题", "行动", "负责人", "截止日期"], ["P0", "华北二店目标缺口与 7.2% 退货率", "复盘退货 Top 20 SKU 并修订补货", "华北区经理", "2026-09-10"], ["P1", "华北整体目标达成率偏低", "制定两周门店辅导计划", "零售运营", "2026-09-15"]],
                    },
                ],
            }
        ],
        "presentations": [
            {
                "filename": outputs[1],
                "slides": [
                    {"title": "门店经营复盘", "metrics": [["总收入", "374 万元"], ["目标达成率", "94.7%"], ["风险门店", "华北二店"]]},
                    {"title": "区域表现", "bullets": ["华东收入 223 万元，超过目标 3 万元。", "华北收入 151 万元，较目标少 24 万元。"]},
                    {"title": "核心风险", "bullets": ["华北二店收入 64 万元，目标达成率 75.3%。", "华北二店退货率 7.2%、满意度 75，风险集中。"]},
                    {"title": "两项优先行动", "table": {"headers": ["行动", "负责人", "截止日期"], "rows": [["复盘退货 Top 20 SKU 并修订补货", "华北区经理", "2026-09-10"], ["制定两周门店辅导计划", "零售运营", "2026-09-15"]]}},
                ],
            }
        ],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {
                "contains": ["华北二店", "行动清单", "7.2%", "零售运营"],
                "exact_terms": ["3740000", "3950000", "640000", "850000"],
                "sheet_names": ["门店明细", "区域汇总", "行动清单"],
                "cells": {"区域汇总!B2": 2230000, "区域汇总!B3": 1510000, "区域汇总!B4": 3740000, "行动清单!D2": "华北区经理"},
                "formulas": {"门店明细!E5": "=C5/D5", "区域汇总!D4": "=B4/C4"},
                "min_charts": 2,
                "excludes": ["{{", "华北二店表现最佳"],
            },
            outputs[1]: {"contains": ["门店经营复盘", "374 万元", "94.7%", "华北二店", "7.2%", "华北区经理"], "exact_terms": ["2026-09-10", "2026-09-15"], "min_slides": 4, "max_slides": 5, "min_shapes": 16, "excludes": ["表现最佳"]},
        },
        "shared_terms": [{"term": "华北二店", "files": outputs}, {"term": "7.2%", "files": outputs}, {"term": "华北区经理", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="门店数据分析与汇报",
        description="从经营明细抽取、分析并交付 Excel 与 PPT。",
        instruction=instruction(
            "完成门店经营分析与管理层复盘",
            "你是零售运营分析师。管理层需要既能下钻复算的 Excel，也需要一份 5 页以内的决策型 PPT。",
            ["门店经营明细.csv", "management_questions.pdf"],
            outputs,
            ["Excel 包含门店明细、区域汇总和行动清单，使用公式计算达成率并放置至少两个图表。", "PPT 回答管理层四个问题，明确风险门店、风险原因和两项行动。", "两个文件对华北二店、退货率和责任人的表述一致。"],
        ),
        metadata=metadata(
            "extract_analyze_present",
            operations=["transform", "analyze", "package"],
            modalities=["spreadsheet", "document", "multi_file"],
            scope="cross_app",
            constraints=["numeric_accuracy", "formula_correctness", "cross_artifact_consistency", "professional_readiness"],
            capabilities=["K1", "K3", "K4", "K5", "K8", "K10", "K13"],
            sector="零售",
            occupation="零售运营分析师",
            outputs=["xlsx", "pptx"],
            difficulty="expert",
            minutes=90,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["cowork"], REFERENCE_URLS["workspace"]],
    )


def build_research_to_deliverables() -> None:
    task_id = "cross-002-research-to-deliverables"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    write_text(
        input_dir / "研究摘要.md",
        """# 研究摘要

- 研究 A：在高频重复咨询中，检索增强可降低幻觉，但必须保留人工升级通道。
- 研究 B：用结构化质检清单评审客服回复，可显著提升事实一致性。
- 研究 C：自动摘要适合生成工单初稿，不适合自动关闭投诉。
""",
    )
    write_csv(input_dir / "pilot_results.csv", [["指标", "试点前", "试点后", "变化"], ["平均处理时长_分钟", 12.5, 10.2, "-18.4%"], ["事实准确率", "86.7%", "92.1%", "+5.4pp"], ["人工升级率", "14.3%", "11.6%", "-2.7pp"]])
    write_json(input_dir / "决策问题.json", {"decision": "是否扩大到全部售后咨询", "guardrail": "投诉不得自动关闭", "next_review": "2026-10-01"})
    outputs = [
        "AI客服试点评估报告.docx",
        "AI客服试点决策汇报.pptx",
        "AI客服试点交互看板.html",
    ]
    solve_spec = {
        "required_inputs": ["研究摘要.md", "pilot_results.csv", "决策问题.json"],
        "documents": [
            {
                "filename": outputs[0],
                "title": "AI 客服试点评估报告",
                "sections": [
                    {"heading": "决策摘要", "paragraphs": ["建议有条件扩大到全部售后咨询，但投诉不得自动关闭，下一次复审日期为 2026-10-01。"]},
                    {"heading": "试点结果", "table": {"headers": ["指标", "试点前", "试点后", "变化"], "rows": [["平均处理时长", "12.5 分钟", "10.2 分钟", "-18.4%"], ["事实准确率", "86.7%", "92.1%", "+5.4pp"], ["人工升级率", "14.3%", "11.6%", "-2.7pp"]]}},
                    {"heading": "研究证据", "bullets": ["研究 A 支持在高频重复咨询中使用检索增强，并保留人工升级通道。", "研究 B 支持引入结构化质检清单。", "研究 C 仅支持生成工单初稿，不支持自动关闭投诉。"]},
                    {"heading": "扩展条件", "bullets": ["投诉类工单强制人工确认。", "每周抽检事实准确率并保持不低于 92.1%。", "在 2026-10-01 复审人工升级率和客户投诉率。"]},
                ],
            }
        ],
        "presentations": [
            {
                "filename": outputs[1],
                "slides": [
                    {"title": "AI 客服试点决策", "metrics": [["处理时长", "-18.4%"], ["事实准确率", "92.1%"], ["人工升级率", "11.6%"]]},
                    {"title": "研究证据怎么支持扩展", "bullets": ["检索增强适合高频重复咨询。", "结构化质检清单提升事实一致性。", "自动摘要只生成工单初稿。"]},
                    {"title": "建议：有条件扩大", "bullets": ["扩大到全部售后咨询。", "投诉不得自动关闭，必须人工确认。", "2026-10-01 复审关键指标。"]},
                ],
            }
        ],
        "html_pages": [
            {
                "filename": outputs[2],
                "title": "AI 客服试点交互看板",
                "subtitle": "试点结果、研究证据与扩展护栏",
                "decision": "建议有条件扩大到全部售后咨询。",
                "guardrail": "投诉不得自动关闭，必须保留人工升级和确认通道。",
                "review_date": "2026-10-01",
                "metrics": [
                    ["平均处理时长变化", "-18.4%"],
                    ["事实准确率", "92.1%"],
                    ["人工升级率", "11.6%"],
                ],
                "evidence": [
                    "研究 A：检索增强适合高频重复咨询，但必须保留人工升级通道。",
                    "研究 B：结构化质检清单可提升事实一致性。",
                    "研究 C：自动摘要只适合生成工单初稿。",
                ],
                "actions": [
                    ["投诉类工单强制人工确认", "客服运营", "立即执行"],
                    ["每周抽检事实准确率", "质检负责人", "每周"],
                    ["复审升级率和投诉率", "产品负责人", "2026-10-01"],
                ],
            }
        ],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["决策摘要", "试点结果", "研究证据", "扩展条件", "有条件扩大", "投诉不得自动关闭"], "exact_terms": ["-18.4%", "92.1%", "11.6%", "2026-10-01"], "min_tables": 1, "min_headings": 4, "min_styled": 5, "excludes": ["全面自动关闭", "{{"]},
            outputs[1]: {"contains": ["AI 客服试点决策", "有条件扩大", "投诉不得自动关闭", "检索增强", "结构化质检清单"], "exact_terms": ["-18.4%", "92.1%", "11.6%", "2026-10-01"], "min_slides": 3, "max_slides": 4, "min_shapes": 12, "excludes": ["全面自动关闭"]},
            outputs[2]: {"contains": ["AI 客服试点交互看板", "有条件扩大", "投诉不得自动关闭", "研究证据", "function toggleEvidence", "aria-expanded"], "exact_terms": ["-18.4%", "92.1%", "11.6%", "2026-10-01"], "min_buttons": 1, "min_scripts": 1, "excludes": ["全面自动关闭", "{{"]},
        },
        "shared_terms": [{"term": "-18.4%", "files": outputs}, {"term": "92.1%", "files": outputs}, {"term": "投诉不得自动关闭", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="研究证据转决策材料",
        description="把研究摘要和试点数据转成报告、决策汇报与交互看板。",
        instruction=instruction(
            "形成 AI 客服试点决策材料",
            "你是 AI 产品策略经理，需要把研究证据和试点结果组合成详细评估报告、管理层决策 PPT 和可离线打开的交互看板。",
            ["研究摘要.md", "pilot_results.csv", "决策问题.json"],
            outputs,
            ["报告清楚区分研究证据、试点事实和建议。", "PPT 在 4 页以内给出明确决策、收益和护栏。", "HTML 看板必须离线可打开，包含可展开/收起的研究证据。", "三个文件中的三项指标、决策、护栏和复审日期一致。"],
        ),
        metadata=metadata(
            "research_to_deliverables",
            operations=["analyze", "create", "package"],
            modalities=["plain_text", "spreadsheet", "multi_file"],
            scope="cross_app",
            constraints=["content_fidelity", "traceability", "cross_artifact_consistency", "professional_readiness"],
            capabilities=["K1", "K2", "K5", "K6", "K8", "K10", "K13"],
            sector="互联网服务",
            occupation="AI 产品策略经理",
            outputs=["docx", "pptx", "html"],
            minutes=80,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["forte"], REFERENCE_URLS["cowork"]],
    )


def build_shared_source_multi_output() -> None:
    task_id = "cross-003-shared-source-multi-output"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    create_xlsx(
        input_dir / "product_launch_master.xlsx",
        [
            {
                "name": "Master",
                "rows": [["字段", "值"], ["产品", "Aurora Air 净化器"], ["上市日期", "2026-10-15"], ["建议零售价", 299], ["首月销量目标", 12000], ["主卖点", "静音 18dB；滤芯寿命 12 个月；适用 35㎡"], ["渠道", "电商/直营门店"]],
            }
        ],
    )
    outputs = ["新品上市主数据.xlsx", "新品上市执行简报.docx", "新品上市沟通会.pptx"]
    solve_spec = {
        "required_inputs": ["product_launch_master.xlsx"],
        "workbooks": [
            {"filename": outputs[0], "sheets": [{"name": "发布主数据", "rows": [["产品", "上市日期", "建议零售价_元", "首月销量目标", "渠道"], ["Aurora Air 净化器", "2026-10-15", 299, 12000, "电商/直营门店"]]}, {"name": "卖点", "rows": [["序号", "卖点"], [1, "静音 18dB"], [2, "滤芯寿命 12 个月"], [3, "适用 35㎡"]]}, {"name": "里程碑", "rows": [["里程碑", "日期", "负责人"], ["物料定稿", "2026-09-25", "品牌团队"], ["渠道培训", "2026-10-08", "销售运营"], ["正式上市", "2026-10-15", "产品负责人"]]}]},
        ],
        "documents": [
            {"filename": outputs[1], "title": "Aurora Air 新品上市执行简报", "sections": [{"heading": "上市目标", "paragraphs": ["Aurora Air 净化器将于 2026-10-15 在电商和直营门店上市，建议零售价 299 元，首月销量目标 12,000 台。"]}, {"heading": "核心卖点", "bullets": ["静音 18dB", "滤芯寿命 12 个月", "适用 35㎡"]}, {"heading": "执行里程碑", "table": {"headers": ["里程碑", "日期", "负责人"], "rows": [["物料定稿", "2026-09-25", "品牌团队"], ["渠道培训", "2026-10-08", "销售运营"], ["正式上市", "2026-10-15", "产品负责人"]]}}]},
        ],
        "presentations": [
            {"filename": outputs[2], "slides": [{"title": "Aurora Air 新品上市沟通会", "metrics": [["上市日期", "2026-10-15"], ["建议零售价", "299 元"], ["首月目标", "12,000 台"]]}, {"title": "为什么值得买", "bullets": ["静音 18dB", "滤芯寿命 12 个月", "适用 35㎡"]}, {"title": "渠道与里程碑", "table": {"headers": ["事项", "日期", "负责人"], "rows": [["物料定稿", "2026-09-25", "品牌团队"], ["渠道培训", "2026-10-08", "销售运营"], ["正式上市", "2026-10-15", "产品负责人"]]}}]},
        ],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["Aurora Air 净化器", "静音 18dB", "滤芯寿命 12 个月", "适用 35㎡"], "sheet_names": ["发布主数据", "卖点", "里程碑"], "cells": {"发布主数据!B2": "2026-10-15", "发布主数据!C2": 299, "发布主数据!D2": 12000}, "excludes": ["{{"]},
            outputs[1]: {"contains": ["Aurora Air 净化器", "上市目标", "核心卖点", "执行里程碑", "电商和直营门店"], "exact_terms": ["2026-10-15", "299 元", "12,000 台"], "min_tables": 1, "min_headings": 3, "min_styled": 4, "excludes": ["{{"]},
            outputs[2]: {"contains": ["Aurora Air 新品上市沟通会", "静音 18dB", "滤芯寿命 12 个月", "适用 35㎡"], "exact_terms": ["2026-10-15", "299 元", "12,000 台"], "min_slides": 3, "max_slides": 4, "min_shapes": 13, "excludes": ["{{"]},
        },
        "shared_terms": [{"term": "2026-10-15", "files": outputs}, {"term": "静音 18dB", "files": outputs}, {"term": "Aurora Air", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="单一事实源多交付物",
        description="从产品主数据生成一致的 Excel、Word 和 PPT 上市包。",
        instruction=instruction(
            "生成新品上市三件套",
            "你是产品上市项目经理。唯一可信事实源是主数据工作簿，请基于它生成可维护的主数据表、执行简报和沟通会 PPT。",
            ["product_launch_master.xlsx"],
            outputs,
            ["Excel 拆分发布主数据、卖点和里程碑。", "Word 说明目标、卖点和执行责任。", "PPT 用于跨团队沟通，4 页以内。", "三个文件中的上市日期、价格、销量目标和卖点完全一致。"],
        ),
        metadata=metadata(
            "shared_source_multi_output",
            operations=["transform", "create", "package"],
            modalities=["spreadsheet"],
            scope="cross_app",
            constraints=["content_fidelity", "cross_artifact_consistency", "style_consistency", "professional_readiness"],
            capabilities=["K1", "K6", "K8", "K9", "K11", "K13"],
            sector="消费电子",
            occupation="产品上市项目经理",
            outputs=["xlsx", "docx", "pptx"],
            minutes=85,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["workspace"], REFERENCE_URLS["cowork"]],
    )


def build_update_propagation() -> None:
    task_id = "cross-004-update-propagation"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    create_xlsx(input_dir / "项目预算_v1.xlsx", [{"name": "预算", "rows": [["项目", "金额_元", "负责人", "上线日期"], ["客户数据平台", 850000, "陈尧", "2026-09-30"]]}])
    create_docx(input_dir / "项目计划_v1.docx", title="客户数据平台项目计划", sections=[{"heading": "关键基线", "paragraphs": ["总预算 850,000 元；负责人陈尧；上线日期 2026-09-30。"]}, {"heading": "范围", "paragraphs": ["客户主数据、标签服务和权限审计。"]}])
    create_pptx(input_dir / "项目状态_v1.pptx", [{"title": "客户数据平台项目状态", "metrics": [["总预算", "850,000 元"], ["负责人", "陈尧"], ["上线日期", "2026-09-30"]]}, {"title": "范围", "bullets": ["客户主数据", "标签服务", "权限审计"]}])
    write_json(input_dir / "变更单.json", {"change_id": "CR-2026-018", "budget": 920000, "owner": "林岚", "launch_date": "2026-10-20", "reason": "新增历史数据迁移与安全复核", "unchanged_scope": ["客户主数据", "标签服务", "权限审计"]})
    outputs = ["项目预算_v2.xlsx", "项目计划_v2.docx", "项目状态_v2.pptx"]
    solve_spec = {
        "required_inputs": ["项目预算_v1.xlsx", "项目计划_v1.docx", "项目状态_v1.pptx", "变更单.json"],
        "workbooks": [
            {
                "filename": outputs[0],
                "sheets": [
                    {
                        "name": "预算",
                        "rows": [
                            ["项目", "金额_元", "负责人", "上线日期", "变更单"],
                            ["客户数据平台", 920000, "林岚", "2026-10-20", "CR-2026-018"],
                        ],
                    },
                    {
                        "name": "变更记录",
                        "rows": [
                            ["变更单", "变更原因", "原预算", "新预算"],
                            ["CR-2026-018", "新增历史数据迁移与安全复核", 850000, 920000],
                        ],
                    },
                ],
            }
        ],
        "documents": [{"filename": outputs[1], "title": "客户数据平台项目计划 v2", "sections": [{"heading": "变更摘要", "paragraphs": ["CR-2026-018：总预算调整为 920,000 元，负责人调整为林岚，上线日期调整为 2026-10-20。原因是新增历史数据迁移与安全复核。"]}, {"heading": "范围（保持不变）", "bullets": ["客户主数据", "标签服务", "权限审计"]}, {"heading": "版本控制", "paragraphs": ["本文件替代 v1，所有后续沟通使用 v2 基线。"]}]}],
        "presentations": [{"filename": outputs[2], "slides": [{"title": "客户数据平台项目状态 v2", "metrics": [["总预算", "920,000 元"], ["负责人", "林岚"], ["上线日期", "2026-10-20"]]}, {"title": "本次变更", "bullets": ["变更单 CR-2026-018", "新增历史数据迁移与安全复核", "原范围保持不变"]}, {"title": "保持不变的范围", "bullets": ["客户主数据", "标签服务", "权限审计"]}]}],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["客户数据平台", "林岚", "CR-2026-018", "新增历史数据迁移与安全复核"], "sheet_names": ["预算", "变更记录"], "cells": {"预算!B2": 920000, "预算!C2": "林岚", "预算!D2": "2026-10-20"}, "preserve_terms": ["客户主数据", "标签服务", "权限审计"], "excludes": ["{{"]},
            outputs[1]: {"contains": ["变更摘要", "范围（保持不变）", "版本控制", "CR-2026-018", "新增历史数据迁移与安全复核"], "exact_terms": ["920,000 元", "林岚", "2026-10-20"], "preserve_terms": ["客户主数据", "标签服务", "权限审计"], "min_headings": 3, "min_styled": 4, "excludes": ["总预算 850,000 元"]},
            outputs[2]: {"contains": ["项目状态 v2", "CR-2026-018", "新增历史数据迁移与安全复核", "客户主数据", "标签服务", "权限审计"], "exact_terms": ["920,000 元", "林岚", "2026-10-20"], "min_slides": 3, "min_shapes": 13, "excludes": ["总预算 850,000 元"]},
        },
        "shared_terms": [{"term": "林岚", "files": outputs}, {"term": "2026-10-20", "files": outputs}, {"term": "CR-2026-018", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="项目基线变更传播",
        description="把同一变更同步到 Excel、Word 和 PPT，并保护未变范围。",
        instruction=instruction(
            "将项目变更同步到三个基线文件",
            "你是 PMO。审批后的变更单是最新事实，请把预算、负责人和上线日期同步到三个项目文件，同时保留未变的范围并留下版本记录。",
            ["项目预算_v1.xlsx", "项目计划_v1.docx", "项目状态_v1.pptx", "变更单.json"],
            outputs,
            ["三个文件都更新为预算 920,000 元、负责人林岚、上线日期 2026-10-20。", "保留客户主数据、标签服务和权限审计三项范围。", "写明变更单编号与原因，旧值只能出现在变更记录中。"],
        ),
        metadata=metadata(
            "update_propagation",
            operations=["edit", "orchestrate"],
            modalities=["spreadsheet", "document", "presentation", "plain_text", "multi_file"],
            scope="cross_app",
            constraints=["preservation", "cross_artifact_consistency", "traceability", "professional_readiness"],
            capabilities=["K7", "K10", "K11", "K13"],
            sector="软件与数据",
            occupation="PMO",
            outputs=["xlsx", "docx", "pptx"],
            difficulty="expert",
            minutes=90,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["workspace"], REFERENCE_URLS["cowork"]],
    )


def build_cross_file_consistency() -> None:
    task_id = "cross-005-cross-file-consistency"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    write_csv(input_dir / "finance_actual.csv", [["期间", "已确认收入_元", "口径"], ["2026-07", 1280000, "财务已开票且满足确认条件"]])
    create_xlsx(input_dir / "sales_forecast.xlsx", [{"name": "Forecast", "rows": [["期间", "预测收入_元", "说明"], ["2026-07", 1350000, "含 70,000 元未验收项目"]]}])
    create_docx(input_dir / "board_draft.docx", title="7 月董事会经营草稿", sections=[{"heading": "收入", "paragraphs": ["7 月收入为 1,350,000 元，已全部确认。"]}, {"heading": "增长", "paragraphs": ["较 6 月增长 12%。"]}])
    write_json(input_dir / "口径规则.json", {"authoritative_actual": "finance_actual.csv", "forecast_use": "只用于解释差异，不得作为已确认收入", "difference_reason": "70,000 元项目尚未验收"})
    outputs = ["收入口径核对.xlsx", "董事会口径差异说明.docx"]
    solve_spec = {
        "required_inputs": ["finance_actual.csv", "sales_forecast.xlsx", "board_draft.docx", "口径规则.json"],
        "workbooks": [{"filename": outputs[0], "sheets": [{"name": "口径核对", "rows": [["来源", "数值_元", "用途", "是否权威实际"], ["finance_actual.csv", 1280000, "已确认收入", "是"], ["sales_forecast.xlsx", 1350000, "预测与差异解释", "否"], ["差异", 70000, "未验收项目", "不计入实际"]]}, {"name": "修正记录", "rows": [["文件", "原表述", "修正表述"], ["board_draft.docx", "1,350,000 元已全部确认", "1,280,000 元已确认；70,000 元待验收"]]}]}],
        "documents": [{"filename": outputs[1], "title": "董事会收入口径差异说明", "sections": [{"heading": "结论", "paragraphs": ["2026 年 7 月权威已确认收入为 1,280,000 元。销售预测为 1,350,000 元，两者相差 70,000 元。"]}, {"heading": "差异原因", "paragraphs": ["70,000 元项目尚未验收，不满足收入确认条件，只能保留在预测口径。"]}, {"heading": "董事会材料修正", "paragraphs": ["将草稿中的“1,350,000 元已全部确认”修正为“1,280,000 元已确认；70,000 元待验收”。"]}, {"heading": "控制建议", "bullets": ["董事会实际收入只引用 finance_actual.csv。", "预测数据必须显式标注为 Forecast。"]}]}],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["口径核对", "修正记录", "已确认收入", "未验收项目"], "sheet_names": ["口径核对", "修正记录"], "cells": {"口径核对!B2": 1280000, "口径核对!B3": 1350000, "口径核对!B4": 70000}, "excludes": ["{{"]},
            outputs[1]: {"contains": ["结论", "差异原因", "董事会材料修正", "控制建议", "finance_actual.csv"], "exact_terms": ["1,280,000 元", "1,350,000 元", "70,000 元"], "min_headings": 4, "min_styled": 5, "excludes": ["1,350,000 元已全部确认。"]},
        },
        "shared_terms": [{"term": "1280000", "files": [outputs[0]]}, {"term": "70,000 元", "files": [outputs[1]]}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="收入口径冲突核对",
        description="识别跨文件收入冲突，按权威来源修正并形成证据链。",
        instruction=instruction(
            "核对并修正董事会收入口径",
            "你是财务 BP。三个文件对 7 月收入表述冲突，请按口径规则确定权威实际、解释差异，并提供可审计核对表与董事会修正说明。",
            ["finance_actual.csv", "sales_forecast.xlsx", "board_draft.docx", "口径规则.json"],
            outputs,
            ["明确 1,280,000 元是权威已确认收入。", "解释 1,350,000 元预测与 70,000 元未验收差异。", "核对表保留来源、用途和修正记录。", "Word 给出董事会材料的准确替换表述。"],
        ),
        metadata=metadata(
            "cross_file_consistency",
            operations=["audit", "analyze", "repair"],
            modalities=["spreadsheet", "document", "plain_text", "multi_file"],
            scope="cross_app",
            constraints=["numeric_accuracy", "traceability", "cross_artifact_consistency", "content_fidelity"],
            capabilities=["K1", "K2", "K5", "K10", "K12"],
            sector="财务",
            occupation="财务 BP",
            outputs=["xlsx", "docx"],
            difficulty="expert",
            minutes=75,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["workspace"], REFERENCE_URLS["officebench"]],
    )


def build_template_driven_package() -> None:
    task_id = "cross-006-template-driven-package"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    create_docx(input_dir / "立项书模板.docx", title="{{项目名}}立项书", sections=[{"heading": "背景", "paragraphs": ["{{背景}}"]}, {"heading": "预算与周期", "paragraphs": ["预算：{{预算}}；启动：{{启动日期}}"]}])
    create_xlsx(input_dir / "预算模板.xlsx", [{"name": "预算", "rows": [["项目", "类别", "金额_元"], ["{{项目名}}", "软件", "{{软件预算}}"], ["{{项目名}}", "实施", "{{实施预算}}"], ["{{项目名}}", "培训", "{{培训预算}}"]]}])
    create_pptx(input_dir / "启动会模板.pptx", [{"title": "{{项目名}}启动会", "metrics": [["总预算", "{{预算}}"], ["启动日期", "{{启动日期}}"], ["负责人", "{{负责人}}"]]}])
    write_json(input_dir / "project_facts.json", {"project": "星港计划", "background": "统一三地仓库的库存与调拨数据", "budget": 2400000, "start_date": "2026-11-01", "owner": "沈舟", "budget_lines": {"软件": 1200000, "实施": 900000, "培训": 300000}, "milestones": [["蓝图确认", "2026-11-15"], ["试点上线", "2027-01-15"], ["三地推广", "2027-03-01"]]})
    outputs = ["星港计划立项书.docx", "星港计划预算.xlsx", "星港计划启动会.pptx"]
    solve_spec = {
        "required_inputs": ["立项书模板.docx", "预算模板.xlsx", "启动会模板.pptx", "project_facts.json"],
        "documents": [{"filename": outputs[0], "title": "星港计划立项书", "sections": [{"heading": "项目背景", "paragraphs": ["星港计划用于统一三地仓库的库存与调拨数据。"]}, {"heading": "预算与启动", "paragraphs": ["总预算 2,400,000 元，启动日期 2026-11-01，项目负责人沈舟。"]}, {"heading": "里程碑", "table": {"headers": ["里程碑", "日期"], "rows": [["蓝图确认", "2026-11-15"], ["试点上线", "2027-01-15"], ["三地推广", "2027-03-01"]]}}, {"heading": "成功标准", "bullets": ["库存与调拨数据采用统一口径。", "三地仓库按里程碑完成推广。"]}]}],
        "workbooks": [{"filename": outputs[1], "sheets": [{"name": "预算", "rows": [["项目", "类别", "金额_元"], ["星港计划", "软件", 1200000], ["星港计划", "实施", 900000], ["星港计划", "培训", 300000], ["星港计划", "合计", None]], "formulas": {"K5": "=SUM(C2:C4)"}}, {"name": "里程碑", "rows": [["里程碑", "日期", "负责人"], ["蓝图确认", "2026-11-15", "沈舟"], ["试点上线", "2027-01-15", "沈舟"], ["三地推广", "2027-03-01", "沈舟"]]}]}],
        "presentations": [{"filename": outputs[2], "slides": [{"title": "星港计划启动会", "metrics": [["总预算", "2,400,000 元"], ["启动日期", "2026-11-01"], ["负责人", "沈舟"]]}, {"title": "目标与范围", "bullets": ["统一三地仓库的库存与调拨数据", "建立统一口径与推广节奏"]}, {"title": "三项里程碑", "table": {"headers": ["里程碑", "日期"], "rows": [["蓝图确认", "2026-11-15"], ["试点上线", "2027-01-15"], ["三地推广", "2027-03-01"]]}}]}],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["星港计划立项书", "项目背景", "预算与启动", "里程碑", "成功标准"], "exact_terms": ["2,400,000 元", "2026-11-01", "沈舟"], "min_tables": 1, "min_headings": 4, "min_styled": 5, "excludes": ["{{", "}}"]},
            outputs[1]: {"contains": ["星港计划", "软件", "实施", "培训", "沈舟"], "sheet_names": ["预算", "里程碑"], "cells": {"预算!C2": 1200000, "预算!C3": 900000, "预算!C4": 300000}, "formulas": {"预算!C5": "=SUM(C2:C4)"}, "excludes": ["{{", "}}"]},
            outputs[2]: {"contains": ["星港计划启动会", "统一三地仓库", "蓝图确认", "试点上线", "三地推广"], "exact_terms": ["2,400,000 元", "2026-11-01", "沈舟"], "min_slides": 3, "min_shapes": 13, "excludes": ["{{", "}}"]},
        },
        "shared_terms": [{"term": "星港计划", "files": outputs}, {"term": "沈舟", "files": outputs}, {"term": "2026-11-01", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="模板驱动项目立项包",
        description="把项目事实准确填入 Word、Excel、PPT 三套模板。",
        instruction=instruction(
            "完成星港计划立项包",
            "你是项目管理办公室专员。请使用公司提供的三种模板和唯一项目事实源，形成可直接审批与启动的完整项目包。",
            ["立项书模板.docx", "预算模板.xlsx", "启动会模板.pptx", "project_facts.json"],
            outputs,
            ["清除全部占位符。", "预算工作簿按软件、实施、培训拆分并用公式求和。", "三份文件的项目名、预算、启动日期、负责人和里程碑一致。", "保留模板所代表的文件类型与专业结构。"],
        ),
        metadata=metadata(
            "template_driven_package",
            operations=["edit", "package", "orchestrate"],
            modalities=["document", "spreadsheet", "presentation", "plain_text", "multi_file"],
            scope="cross_app",
            constraints=["template_adherence" if False else "style_consistency", "numeric_accuracy", "cross_artifact_consistency", "professional_readiness"],
            capabilities=["K4", "K7", "K9", "K11", "K13"],
            sector="物流与供应链",
            occupation="PMO 专员",
            outputs=["docx", "xlsx", "pptx"],
            difficulty="expert",
            minutes=95,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["officeval"], REFERENCE_URLS["workspace"]],
    )


def build_review_and_revision() -> None:
    task_id = "cross-007-review-and-revision"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    create_pptx(input_dir / "月度产品复盘_草稿.pptx", [{"title": "月度产品复盘", "metrics": [["月活用户", "136,000"], ["流失率", "3.2%"], ["NPS", "48"]]}, {"title": "结论", "bullets": ["增长健康，无需额外动作。"]}])
    create_docx(input_dir / "评审意见.docx", title="月度产品复盘评审意见", sections=[{"heading": "数据问题", "bullets": ["月活用户应为 128,000，不是 136,000。", "流失率应为 3.8%，不是 3.2%。"]}, {"heading": "叙事问题", "bullets": ["必须解释流失率上升。", "增加数据口径页。", "把详细渠道表移到附录。"]}])
    create_xlsx(
        input_dir / "verified_metrics.xlsx",
        [
            {
                "name": "Metrics",
                "rows": [
                    ["指标", "数值", "口径"],
                    ["月活用户", 128000, "月末 30 天去重活跃"],
                    ["流失率", "3.8%", "月初存量用户流失"],
                    ["NPS", 48, "有效问卷 1,204 份"],
                ],
            },
            {
                "name": "Channels",
                "rows": [
                    ["渠道", "新增用户"],
                    ["自然流量", 22000],
                    ["付费投放", 18000],
                    ["合作渠道", 9000],
                ],
            },
        ],
    )
    outputs = ["月度产品复盘_修订版.pptx", "评审意见响应表.docx"]
    solve_spec = {
        "required_inputs": ["月度产品复盘_草稿.pptx", "评审意见.docx", "verified_metrics.xlsx"],
        "presentations": [{"filename": outputs[0], "slides": [{"title": "月度产品复盘（修订版）", "metrics": [["月活用户", "128,000"], ["流失率", "3.8%"], ["NPS", "48"]]}, {"title": "核心诊断", "bullets": ["月活用户为 128,000。", "流失率升至 3.8%，需要加强新用户首周激活与召回。"]}, {"title": "下一步行动", "table": {"headers": ["行动", "负责人", "时间"], "rows": [["分析首周未激活用户", "增长产品", "2026-09-05"], ["上线分层召回实验", "用户运营", "2026-09-12"]]}}, {"title": "数据口径", "bullets": ["月活用户：月末 30 天去重活跃。", "流失率：月初存量用户流失。", "NPS：有效问卷 1,204 份。"]}, {"title": "附录：渠道新增", "table": {"headers": ["渠道", "新增用户"], "rows": [["自然流量", "22,000"], ["付费投放", "18,000"], ["合作渠道", "9,000"]]}}]}],
        "documents": [{"filename": outputs[1], "title": "评审意见响应表", "sections": [{"heading": "响应总览", "paragraphs": ["已逐项处理数据更正、流失率解释、数据口径页和渠道附录四类意见。"]}, {"heading": "逐项响应", "table": {"headers": ["意见", "处理结果", "修订位置"], "rows": [["月活用户更正为 128,000", "已完成", "第 1-2 页"], ["流失率更正为 3.8%", "已完成并增加诊断", "第 1-2 页"], ["增加数据口径页", "已完成", "第 4 页"], ["渠道表移至附录", "已完成", "第 5 页"]]}}, {"heading": "保留项", "paragraphs": ["NPS 48 与有效问卷 1,204 份保持不变。"]}]}],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["月度产品复盘（修订版）", "核心诊断", "下一步行动", "数据口径", "附录：渠道新增", "月末 30 天去重活跃"], "exact_terms": ["128,000", "3.8%", "NPS", "48", "1,204"], "min_slides": 5, "max_slides": 6, "min_shapes": 20, "excludes": ["136,000", "3.2%", "无需额外动作"]},
            outputs[1]: {"contains": ["响应总览", "逐项响应", "保留项", "已完成", "第 4 页", "第 5 页"], "exact_terms": ["128,000", "3.8%", "NPS 48", "1,204"], "min_tables": 1, "min_headings": 3, "min_styled": 4, "excludes": ["未处理"]},
        },
        "shared_terms": [{"term": "128,000", "files": outputs}, {"term": "3.8%", "files": outputs}, {"term": "1,204", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="基于评审意见修订复盘",
        description="联合评审意见与权威数据修订 PPT，并提交响应表。",
        instruction=instruction(
            "修订月度产品复盘并回应评审",
            "你是产品运营经理。草稿中有错误数据和不完整叙事，请以 verified_metrics.xlsx 为权威来源逐项处理 Word 评审意见。",
            ["月度产品复盘_草稿.pptx", "评审意见.docx", "verified_metrics.xlsx"],
            outputs,
            ["把月活和流失率改为权威值，保留正确的 NPS。", "新增流失诊断、行动、数据口径页，并把渠道表移到附录。", "响应表逐项写明处理结果和修订位置。", "两个文件中的指标和页码引用一致。"],
        ),
        metadata=metadata(
            "review_and_revision",
            operations=["audit", "edit", "repair", "package"],
            modalities=["presentation", "document", "spreadsheet", "multi_file"],
            scope="cross_app",
            constraints=["content_fidelity", "numeric_accuracy", "preservation", "traceability", "cross_artifact_consistency"],
            capabilities=["K2", "K5", "K7", "K8", "K10", "K12", "K13"],
            sector="互联网产品",
            occupation="产品运营经理",
            outputs=["pptx", "docx"],
            difficulty="expert",
            minutes=85,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["officeval"], REFERENCE_URLS["cowork"]],
    )


def build_workspace_orchestration() -> None:
    task_id = "cross-008-workspace-orchestration"
    root = TASKS_ROOT / task_id
    input_dir = root / "data/input_files"
    create_docx(input_dir / "brief/项目简报.docx", title="海风计划项目简报", sections=[{"heading": "目标", "paragraphs": ["在 2026-12-01 上线新的会员数据平台。"]}, {"heading": "负责人", "paragraphs": ["项目负责人：顾清；数据迁移负责人：魏然。"]}, {"heading": "范围", "bullets": ["会员主数据", "积分迁移", "权限审计"]}])
    create_xlsx(input_dir / "finance/预算上限.xlsx", [{"name": "Cap", "rows": [["项目", "预算上限_元", "已承诺_元"], ["海风计划", 1500000, 980000]]}])
    write_csv(input_dir / "planning/milestones.csv", [["里程碑", "日期", "负责人"], ["数据冻结", "2026-10-15", "魏然"], ["迁移演练", "2026-11-05", "魏然"], ["业务验收", "2026-11-20", "顾清"], ["正式上线", "2026-12-01", "顾清"]])
    write_text(input_dir / "meetings/评审纪要.md", """# 评审纪要

- P0 风险：历史会员 ID 存在重复，迁移前必须完成去重规则确认。
- 预算剩余额度只能用于迁移演练和安全复核。
- 启动会需要明确四个里程碑、两个负责人和升级路径。
""")
    outputs = ["海风计划执行计划.xlsx", "海风计划启动会.pptx", "海风计划行动备忘录.docx"]
    solve_spec = {
        "required_inputs": ["brief/项目简报.docx", "finance/预算上限.xlsx", "planning/milestones.csv", "meetings/评审纪要.md"],
        "workbooks": [{"filename": outputs[0], "sheets": [{"name": "里程碑", "rows": [["里程碑", "日期", "负责人", "状态"], ["数据冻结", "2026-10-15", "魏然", "未开始"], ["迁移演练", "2026-11-05", "魏然", "未开始"], ["业务验收", "2026-11-20", "顾清", "未开始"], ["正式上线", "2026-12-01", "顾清", "未开始"]]}, {"name": "预算", "rows": [["项目", "预算上限_元", "已承诺_元", "剩余额度_元"], ["海风计划", 1500000, 980000, None]], "formulas": {"D2": "=B2-C2"}}, {"name": "风险与行动", "rows": [["级别", "风险", "行动", "负责人", "截止日期"], ["P0", "历史会员 ID 重复", "确认去重规则并完成样本验证", "魏然", "2026-10-10"], ["P1", "迁移演练与安全复核资金受限", "按剩余额度排定采购", "顾清", "2026-10-20"]]}]}],
        "presentations": [{"filename": outputs[1], "slides": [{"title": "海风计划启动会", "metrics": [["预算上限", "1,500,000 元"], ["上线日期", "2026-12-01"], ["负责人", "顾清"]]}, {"title": "范围与角色", "bullets": ["会员主数据、积分迁移、权限审计", "项目负责人顾清；数据迁移负责人魏然"]}, {"title": "四个里程碑", "table": {"headers": ["里程碑", "日期", "负责人"], "rows": [["数据冻结", "2026-10-15", "魏然"], ["迁移演练", "2026-11-05", "魏然"], ["业务验收", "2026-11-20", "顾清"], ["正式上线", "2026-12-01", "顾清"]]}}, {"title": "P0 风险与升级路径", "bullets": ["历史会员 ID 重复，迁移前必须确认去重规则。", "魏然负责样本验证；逾期升级给顾清。"]}]}],
        "documents": [{"filename": outputs[2], "title": "海风计划行动备忘录", "sections": [{"heading": "项目基线", "paragraphs": ["预算上限 1,500,000 元，已承诺 980,000 元，剩余额度 520,000 元；正式上线日期 2026-12-01。"]}, {"heading": "P0 行动", "paragraphs": ["魏然须在 2026-10-10 前确认历史会员 ID 去重规则并完成样本验证；逾期升级给项目负责人顾清。"]}, {"heading": "里程碑", "table": {"headers": ["里程碑", "日期", "负责人"], "rows": [["数据冻结", "2026-10-15", "魏然"], ["迁移演练", "2026-11-05", "魏然"], ["业务验收", "2026-11-20", "顾清"], ["正式上线", "2026-12-01", "顾清"]]}}, {"heading": "预算约束", "paragraphs": ["剩余额度仅用于迁移演练和安全复核。"]}]}],
    }
    expectations = {
        "required_files": outputs,
        "files": {
            outputs[0]: {"contains": ["海风计划", "历史会员 ID 重复", "魏然", "顾清", "风险与行动"], "sheet_names": ["里程碑", "预算", "风险与行动"], "cells": {"预算!B2": 1500000, "预算!C2": 980000, "风险与行动!D2": "魏然"}, "formulas": {"预算!D2": "=B2-C2"}, "excludes": ["{{"]},
            outputs[1]: {"contains": ["海风计划启动会", "会员主数据", "积分迁移", "权限审计", "历史会员 ID 重复", "逾期升级给顾清"], "exact_terms": ["1,500,000 元", "2026-12-01", "魏然"], "min_slides": 4, "min_shapes": 17, "excludes": ["{{"]},
            outputs[2]: {"contains": ["项目基线", "P0 行动", "里程碑", "预算约束", "历史会员 ID 去重规则", "迁移演练和安全复核"], "exact_terms": ["1,500,000 元", "980,000 元", "520,000 元", "2026-12-01", "2026-10-10"], "min_tables": 1, "min_headings": 4, "min_styled": 5, "excludes": ["{{"]},
        },
        "shared_terms": [{"term": "海风计划", "files": outputs}, {"term": "2026-12-01", "files": outputs}, {"term": "魏然", "files": outputs}, {"term": "顾清", "files": outputs}],
    }
    write_task_package(
        root,
        task_id=task_id,
        title="复杂工作区项目编排",
        description="探索分层工作区并完成计划、启动会和行动备忘录。",
        instruction=instruction(
            "编排海风计划启动交付",
            "你是交付项目经理。输入分散在 brief、finance、planning、meetings 四个目录中，请先建立事实基线，再交付可执行计划、启动会和行动备忘录。",
            ["brief/项目简报.docx", "finance/预算上限.xlsx", "planning/milestones.csv", "meetings/评审纪要.md"],
            outputs,
            ["Excel 统一里程碑、预算公式、风险和责任。", "PPT 说明项目基线、范围、四个里程碑和 P0 升级路径。", "Word 汇总必须立即执行的行动与预算约束。", "三个文件中的日期、预算、负责人和 P0 风险一致。"],
        ),
        metadata=metadata(
            "workspace_orchestration",
            operations=["analyze", "package", "orchestrate"],
            modalities=["workspace", "document", "spreadsheet", "plain_text", "multi_file"],
            scope="workspace",
            constraints=["numeric_accuracy", "traceability", "cross_artifact_consistency", "professional_readiness"],
            capabilities=["K1", "K2", "K4", "K5", "K10", "K11", "K13"],
            sector="零售科技",
            occupation="交付项目经理",
            outputs=["xlsx", "pptx", "docx"],
            difficulty="expert",
            minutes=105,
        ),
        solve_spec=solve_spec,
        expectations=expectations,
        source_urls=[REFERENCE_URLS["workspace"], REFERENCE_URLS["cowork"], REFERENCE_URLS["officebench"]],
    )


def main() -> None:
    builders = [
        build_extract_analyze_present,
        build_research_to_deliverables,
        build_shared_source_multi_output,
        build_update_propagation,
        build_cross_file_consistency,
        build_template_driven_package,
        build_review_and_revision,
        build_workspace_orchestration,
    ]
    for builder in builders:
        builder()
        print(f"已生成：{builder.__name__}")
    print(f"共生成 {len(builders)} 个跨 Office 公开端到端样例。")


if __name__ == "__main__":
    main()
