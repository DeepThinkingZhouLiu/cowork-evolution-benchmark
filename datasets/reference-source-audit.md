# Excel 公开题源审计（2026-08-28）

## 范围与结论

本轮只准备上游候选题，不生成 Harbor 任务包，也不使用制题 Skill。现有 14 道 Excel 题保留在 `xlsx-baseline-audit.csv`，不计入本轮候选。

逐题清单见 `xlsx-candidate-inventory.csv`。清单含 97 道新候选：

| 可用性 | 数量 | 含义 |
|---|---:|---|
| `ready` | 70 | 已找到具体任务、输入依据、真值或检查器以及可用许可；进入制题前仍要下载、哈希、打开并逐文件核权 |
| `conditional` | 27 | 任务本身可用，但有明确门槛：数据资产许可、非商业限制、运行时附件缺失或离线重建 |

初始难度分布为 L1=9、L2=36、L3=37、L4=15。这只是结构分层，不是最终难度标签：

- L1：单应用基础操作。
- L2：单应用复杂操作。
- L3：跨应用操作。
- L4：跨应用长程或强状态任务。
- 专业复杂度例外：金融模型、FIFO、复杂薪酬预测等即使只使用 Excel，也可先放入 L3/L4。
- 最终等级必须由无辅助的 Codex/OpenCode rollout 结果和人工复核决定。

## 入选来源

| 来源 | 代码/数据版本 | 许可结论 | 入选 | 源级判断 |
|---|---|---|---:|---|
| OfficeBench | commit `b978b808667c32b52ce19a67ce1def1de9ae02b7` | Apache-2.0 | 24 | 有具体任务 JSON、XLSX 资产/参考文件和确定性 evaluator；部分旧 evaluator 只查文件存在，制题前需增强真值 |
| OmegaUse-OfficeVal | repo `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`; HF `cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5` | Apache-2.0 | 14 | 真实办公请求经隐私改写；有固定附件、人工工时、细粒度 rubric 和逐题 verifier |
| SpreadsheetBench-2 sample | code `5c160265aa93c15b38e4034cbf1e09ab498335d9`; HF `d8a14da1898e48d07c8a9ff8974d0effc9302ece` | sample 数据 CC-BY-SA-4.0 | 12 | 每题有 input XLSX、Golden XLSX、目标范围；ShareAlike 要求必须进入 provenance 和发布策略 |
| Spreadsheet-RL | HF `6f274cd4cbc27e656d7ddae7c3b591669866dc84` | CC-BY-SA-4.0 | 12 | 只从 5,925 道 train 题抽样；parquet 提供任务、目标范围、ground-truth id 和难度元数据，工作簿在归档中 |
| Workspace-Bench-Lite | repo `3fbd0f1a136720fece86786545983e26642c3db2`; HF `60b08b1cc2e8054afbc3ca2160d37876b4f0765c` | 代码 MIT；HF 数据未声明许可证 | 10 | 任务、data manifest、依赖图和细粒度 rubric 完整，但数据许可未明确，全部为 `conditional` |
| BlueFin | commit `643336ba21526666b51bc9e2decd25d57d0e3434` | CC-BY-NC-4.0 | 8 | 公开 7 道 manipulation 和 1 道 synthesis；有 input、指令、rubric 与扰动式 Judge；商业使用被禁止 |
| TheAgentCompany | commit `98b68ef82a47690c316f42fddb05baafaab56851` | MIT | 7 | 任务和 evaluator 完整；2 道输入已入库，5 道依赖 OwnCloud/容器附件，需提取并单独核权 |
| AutomationBench | commit `4a8e1061254004d9dac807054eed33fad7d1ff14` | MIT；许可明确覆盖 mock 数据生成器 | 4 | 无 XLSX 文件，但任务代码内嵌完整表格状态和 assertions；只允许标成离线重建，不能冒充现成工作簿任务 |
| Cowork-Bench | commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d` | Apache-2.0 | 6 | initial workspace、Golden、expected JSON 和 evaluator 完整；适合 L3/L4，但动态 API/数据库需固定快照 |

Spreadsheet-RL 的 12 道全部来自 `train`，没有使用 `test_verified` 或 `test_domain`。SpreadsheetBench-2 已落地的 `Debugging/09_01` 属于现有 14 题基线，本轮没有重复选入。

## 条件项的发布阻断

以下问题在解决前不得把对应候选改为 `ready`：

| 来源 | 阻断条件 |
|---|---|
| Workspace-Bench-Lite | 数据仓没有 license tag；代码仓 MIT 不自动覆盖 HF 中的 DOCX/XLSX/PDF/音频等资产 |
| BlueFin | CC-BY-NC-4.0 禁止商业使用；需先确定 CoworkEvoBench 的发布和使用边界 |
| TheAgentCompany | OwnCloud/容器镜像中的输入没有逐文件许可记录；仅仓库内 `task.md` 和 MIT 代码不足以授权附件再发布 |
| AutomationBench | 只有模拟 Google Sheets/Gmail 状态，没有上游 XLSX；需确定性重建 XLSX/EML 并标记 `offline_reconstruction` 与 `synthetic=true` |
| OfficeBench 弱检查题 | `2-19/0`、`3-99/0` 等 evaluator 只检查文件存在或少量关键词；必须补数值/结构/跨文件一致性真值 |

## 未进入 97 道池的已审计来源

| 来源 | 固定版本与许可 | 本轮结论 |
|---|---|---|
| STATE-Bench | `5644b1838d96bc4483da29642d058ecaa6f80f7f`, MIT | 没有直接 Excel 输入/交付资产；只参考状态变化、经验学习和 held-out 设计 |
| WorkArena++ | `a772230a94cf1caf4166b8ead3983f3b3786455b`, Apache-2.0 | ServiceNow 状态流程，无直接 Excel 题包；只参考组合工作流与检查点 |
| AppWorld | `a072b7a86e7c1d5b1d7175659d750ebb9b79f10a`, Apache-2.0 | API 世界和 state-based evaluation，无直接 Excel 资产 |
| WorkBench | `49c7dfd00c03d384ec59ea57374f50b766aa5613`, MIT | 数据库/API 工作任务，无直接 Office artifact；只参考副作用检查 |
| OSWorld 2.0 | `597c04976738b20e47267b55a0c76bfe0c815eb4`, Apache-2.0 | 当前审计树没有可直接采用的正式 2.0 XLSX 题资产；保留为 GUI 隔离测试来源 |
| ClawMark | `d1b641b3171e584e69a3763c269069f32a13b574`, CC-BY-NC-4.0 | 有动态 CSV/文件任务，但本轮没有发现比已选来源更完整的 Excel 输入+真值组合；保留为 L4 动态注入方法参考 |
| SpreadsheetBench | HF `ab0b742b0fc95b946f212d80ac7771b5531272e4`, CC-BY-SA-4.0 | 与 Spreadsheet-RL/SpreadsheetBench-2 家族重叠；首批不再抽样，避免同源泄漏 |
| SpreadsheetBench-2 full | HF `9dea60025792fbac5928ce9f44812362dccbeecd`, 数据卡 MIT tag | 访问受限且许可与 sample 的 CC-BY-SA 声明不一致；本轮只使用公开 sample 固定 revision |

## 明确排除

- APEX-Agents：数据卡限制训练、自动处理和抓取，只能作为 sealed external evaluation；不进入候选池。
- GDPval：未确认允许衍生和再发布；不进入制题数据源。
- OfficeQA / OfficeComprehensionBench：主要是理解问答，不是直接 Excel 编辑任务；只可另行改写并重新核权。
- 现有 14 道 Excel 题：已经人工审计，是基线而非本轮候选。

## 逐题进入制题阶段的门槛

每道候选只有满足以下条件才可交给制题 Skill：

1. 固定上游任务文本、代码 commit、数据 revision、下载日期和 URL。
2. 实际下载并打开所有输入；记录 SHA256、工作表、隐藏对象、外链、宏、批注、图片和损坏情况。
3. 确认仓库许可与单项资产许可；保留许可证副本、署名和 ShareAlike/NonCommercial 条款。
4. 验证 Golden、rubric 或 evaluator 能区分未作答、部分作答、错误作答和损坏文件。
5. 记录改写边界、`synthetic`、`human_rewritten`、是否允许训练、是否只允许评测。
6. 按 `leakage_group` 做家族级隔离，避免同工作簿、同模板或同 Golden 跨 D-evo/D-test。
7. 制题完成后再执行无辅助 Codex/OpenCode rollout；本表的 L1-L4 不直接成为正式标签。
