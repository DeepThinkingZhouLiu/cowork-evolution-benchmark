# 公开参考源审计

## 审计目的

本文件记录 `003-Reference/002-Reference-Code` 下公开仓库的精确版本、许可证和当前用途。它回答“能否用来挑题”，不直接回答“某个二进制输入能否再发布”。每道候选题仍需单独检查上游任务、附件来源、隐私和资产许可证。

## 当前结论

- 优先题源是 OmegaUse-OfficeVal：它提供 100 条公开办公任务的原始 Query、输入文件链接、Rubric 和逐题 Verifier，并说明任务来自经过隐私保护改写和输入重建的真实请求。
- Cowork-Bench 有 497 个多工具工作区任务，适合补跨 Office 家族；但不少题依赖模拟 API、数据库、邮件或日历，必须先固化为静态、离线、自包含输入，不能直接复制目录就算改造完成。
- LLM-DOCX-Editing 有 50 个 Word 编辑测试和 34 个 DOCX fixture，适合精确编辑、格式保持、重写和红线能力；它更像能力测试，不应包装成“真实用户请求”。
- 没有根目录许可证的仓库只用于方法参考，不复制其代码或 README。如果同一项目的独立数据集明确声明许可证，只可在固定数据集 revision 后逐文件使用该部分资产。
- AGPL-3.0 仓库需要单独评估衍生与分发义务，当前不作为首批资产来源。
- `synthetic` 采用保守口径：只要上游输入经过重建、静态化或属于公开 Benchmark fixture，就标为 `true`，不把它包装成未经处理的企业生产数据。因此当前 43 道候选题全部为 `synthetic=true`。这与“本项目自编题”不是一回事；每题仍必须保存固定上游版本、原 Query/任务记录、输入哈希和适配边界。
- OmegaUse-OfficeVal 数据卡明确说明 100 道任务源自真实办公请求，但经过隐私保护的指令改写、专家复核、输入文件重建和人工脱敏。因此本项目采用的 22 道 OfficeVal 任务统一标为 `human_rewritten=true`、`synthetic=true`。

## 仓库级记录

| 仓库                       | 精确 commit                               | 根许可证     | 可用内容与当前处理                                      |
|----------------------------|-------------------------------------------|--------------|---------------------------------------------------------|
| AutomationBench            | `4a8e1061254004d9dac807054eed33fad7d1ff14` | 未发现       | 财务、HR、营销等多应用状态任务；已筛出 15 个 L3/L4 Excel 工作流 Seed，但只可重建离线输入与业务逻辑，不直接复制无许可资产 |
| Cowork-Bench               | `d943e75bc0fc8e3b27141979300cd8cbcd1e890d` | Apache-2.0   | 497 个工作区任务；已逐题静态化 15 道候选题，包括 2 道演示、1 道文档、6 道表格和 6 道跨办公应用题 |
| DOCX-Redlining-Benchmark   | `17382b6d8916397b4df8db6d2de7f1df6a130fe9` | AGPL-3.0     | 红线方法参考；首批不复制资产                            |
| DeckBench                  | `47393cf4cb8590215ec2276b9603af05e90015f4` | 未发现       | 演示生成/多轮编辑方法参考；资产不进入任务包             |
| EnterpriseClawBench        | `68507699d1a61700f4281f1f34dcadb20f2de06d` | 未发现       | 企业 Agent 设计参考；资产不进入任务包                   |
| FORTE                      | `345c1ec1487139db9dd319787fa9405ba85d1869` | MIT          | 15 个带输入和 Rubric 的职业 demo；已逐题采用 `Legal-020` 和 `Finance-018`，原 Query、输入、Rubric 与上游 Golden 均保留 |
| LLM-DOCX-Editing           | `498a49465d95fce7c0a0b61ac2da6dbd70629d0e` | MIT          | 50 个 Word 编辑测试和 34 个 fixture；能力题候选          |
| OfficeBench                | `b978b808667c32b52ce19a67ce1def1de9ae02b7` | Apache-2.0   | 发现 61 个带 XLSX 输入的任务目录，首轮筛出 30 个单应用/跨应用具体任务；逐题附件许可仍待复核 |
| STATE-Bench                | `5644b1838d96bc4483da29642d058ecaa6f80f7f` | MIT          | 状态转移、冲突信息与 held-out 轨迹参考；无直接 Excel 交付资产，只用于 L3 任务构造与 state-based Judge 设计 |
| OSWorld 2.0                | `597c04976738b20e47267b55a0c76bfe0c815eb4` | Apache-2.0   | 仓库实际包含 47 个 `libreoffice_calc` 示例；首轮选出 6 个具体 ID 审计，必须再核对单题输入许可与是否属于 2.0 正式子集 |
| ClawMark                   | `d1b641b3171e584e69a3763c269069f32a13b574` | 未发现       | 多日动态任务，含中途注入的 XLSX/CSV/邮件/多模态资产；已选 6 个 L4 工作流 Seed，许可澄清前不复制资产 |
| TheAgentCompany            | `98b68ef82a47690c316f42fddb05baafaab56851` | MIT          | 175 道企业工作区任务；已筛出 15 道表格、预算、对账、发票与跨应用任务，待逐题检查 ownCloud 内附件许可 |
| WorkArena++                | `a772230a94cf1caf4166b8ead3983f3b3786455b` | 未发现       | ServiceNow 组合式流程和隐藏状态参考；无直接 Excel 交付资产，只用于 L3/L4 工作流和检查点设计 |
| AppWorld                   | `a072b7a86e7c1d5b1d7175659d750ebb9b79f10a` | Apache-2.0   | 多应用 API 状态变更和 state-based evaluation 参考；任务数据为 bundle，无直接 Excel 资产，不作首轮可发布 Seed |
| OmegaUse-OfficeVal         | `ffbeecb8752447c8e40b594a0eeb1db7236ecb36` | Apache-2.0   | 100 条隐私改写请求、重建输入、Rubric、Verifier；已逐题采用 22 道，其中 `officeval_080` 还固定到 HF data commit `cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5` |
| PPT-Eval                   | `1b8b55a29e48fdc65d423689b6f2370ad91beeea` | MIT          | 演示文稿评测器参考，不直接充当任务数据                   |
| PPTAgent                   | `2419d30b134a71486523e95ded60b32489fd3c61` | MIT          | PDF/PPTX 案例和生成方法；第三方附件需另审版权            |
| PPTArena                   | `6f26a63ec1014b98c44990c7b58133b832070be2` | 未发现       | PPT 编辑评测方法参考；资产不进入任务包                   |
| PPTC                       | `277f56ffc548ce51024d5cc7c0607aeb49b584b8` | MIT          | PPT 生成工具和方法参考                                   |
| UltraPresent              | Hugging Face train row `1381`              | 未声明       | 数据集卡未标注整库许可；所选 arXiv `2407.18219` 附件自身为 CC BY 4.0，只作许可待确认候选 |
| SpreadsheetBench           | `f9ae2bb897da711e4d4cd36e1f78be0c94cfd5a6` | 未发现       | 表格任务定义参考；正式数据许可证确认前不复制             |
| SpreadsheetBench-2         | `5c160265aa93c15b38e4034cbf1e09ab498335d9` | 代码仓未发现 | 代码仓只作方法审计，不复制其文件；单独的 `icyCreater/SpreadsheetBench-2-sample` 数据集明确为 CC-BY-SA-4.0，已固定 revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece` 并逐格改造 `Debugging/09_01` |
| WorkBench                  | `49c7dfd00c03d384ec59ea57374f50b766aa5613` | MIT          | 工作任务研究方法参考                                     |
| Workspace-Bench            | `a3a2ee47890f0bc6a0509b1f71df7183563fc43f` | MIT          | 工作区协议、恢复和评分方法参考                           |

## Excel 候选池快照（2026-08-28）

`xlsx-candidate-inventory.csv` 当前记录 89 条：14 道现有 Excel 题、72 条可继续逐题审计的直接或工作流 Seed，以及 3 条只用于结构/Judge 设计的参考项。状态分布为 `adopted=14`、`source_screened=57`、`workflow_screened=15`、`reference_only=3`。

| 来源 | 条数 | 当前用途 |
|---|---:|---|
| OfficeBench | 30 | 单应用、双应用和三应用任务 Seed；逐题检查 XLSX/EML/PDF/ICS 资产 |
| AutomationBench | 15 | L3 财务、HR、营销工作流；因根目录未发现许可证，只重建逻辑与静态输入 |
| TheAgentCompany | 15 | L3/L4 表格、预算、发票、对账和跨应用任务；ownCloud 附件待逐题核权 |
| OSWorld 2.0 | 6 | Calc 任务 Seed；核对是否属于正式 2.0 子集及单题附件 |
| ClawMark | 6 | L4 动态注入和多日工作流；许可证明确前不复制资产 |
| Cowork-Bench / OmegaUse-OfficeVal / FORTE / SpreadsheetBench-2 | 14 | 已落地 Excel 题，作为现有基线和改造范例 |
| STATE-Bench / WorkArena++ / AppWorld | 3 | 仅参考状态变化、组合流程与 state-based evaluation，不计为可直接发布的 Excel 资产 |

作者初评分布为 `L1=9`、`L2=19`、`L3=43`、`L4=17`、`L3-L4=1`。这些值只用于候选池配额和抽样，必须经 Agent Rollout 与双人审核后才能成为校准等级。

## 单题采用规则

一条上游任务只有在以下内容全部核对后才能进入 `tasks/`：

- 原始 Query、输入文件、Rubric/Verifier 或等价验收依据已经逐项阅读。
- 输入文件实际下载并打开过，截图、工作表、页面、批注、备注、媒体和隐藏对象已经按任务需要检查。
- 上游仓库许可证和单个附件的来源没有明显冲突；许可证副本随题保留。
- 输入 SHA256、上游文本 SHA256、精确 commit、下载时间和来源 URL 写入 provenance。
- 改造记录明确区分上游原文、本项目补充的 Harbor 边界和新增 Judge 判据。
- Golden 不是由通用生成器批量制造，而是围绕该题单独实现并经人工视觉检查。
- 原始未作答文件、损坏文件和至少一个真实部分完成样本用于校准 Judge。

仓库级审计结论会随着单题检查更新；任何“许可证待确认”状态默认阻止公开发布。
