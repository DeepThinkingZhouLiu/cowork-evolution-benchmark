# Excel 60 题重分类与补题计划

## 目的与边界

当前 `tasks/xlsx` 工作目录有 60 个包，但其中 5 道 Cross Office 和 1 道 Word 包暂由学长负责迁移。因此，当前已完成 Excel 交付合同整理的是 54 道：它们的求解侧交付均为一个原生 XLSX；输入可以是多文件、多模态、多跳，复杂度也可以很高。异质的多个输出共同构成成果时，任务属于 `cross_office`。本轮不移动或改造这 6 个非 Excel 包。

本计划只重整领域和交付物边界，不把原有 rollout 分数用来决定题目属于哪个领域。任务难度仍先按工作流结构标注，再由独立 rollout 校准。

## 待写入分类法的判定规则

| 情形 | 领域 | 处理 |
|---|---|---|
| 一个可编辑 XLSX 是唯一 solver 交付；输入可含 DOCX、PDF、JSON、CSV、图片或多个工作簿 | `xlsx` | 保留或改造成单一 XLSX。 |
| 多个 XLSX 被要求分别交付 | `xlsx` 的例外待消除 | 优先合并为一个工作簿的多个工作表；不因同类型文件数量自动变为 Cross Office。 |
| 两种或以上异质输出共同完成目标，且 Judge 必须检查它们之间的一致性，例如 XLSX + DOCX、PPTX、EML、ICS、JSON 或处理器 | `cross_office` | 保留原工作流，迁移任务包和元数据。 |
| 唯一交付是 DOCX | `docx` | 迁移到 Word，输入是否为 XLSX 不影响领域。 |
| 唯一交付是 TXT/ICS，且不存在可自然保留的单一 Office 主交付 | 不保留在 `xlsx` | 归档并用有来源的新 XLSX 候选替换；不得为了凑数凭空增加第二种输出。 |

`cross_office` 的定义应从“至少两个 Office 交付物”改为“至少两个异质、相互依赖的交付物”，以覆盖邮件和日历等真实工作区产物。单一领域默认要求一个主文件；同类型多文件必须有明确例外说明，并优先合并。

## 审计事实与总量

当前 `metadata/task.json` 的输出声明为：45 道单一 XLSX、1 道两个 XLSX、10 道 XLSX 加异质输出、4 道无 XLSX 的输出。按 `instruction.md` 的路径和文字审计得到的是 40 / 14 / 6，二者不能直接互换。

在任何迁移或改造前，建立 `datasets/xlsx-output-contract-audit.csv`。每题一行，至少记录：

```text
task_id, instruction_outputs, task_toml_outputs, metadata_outputs,
expectations_outputs, judge_required_outputs, golden_outputs,
classification_decision, change_type, rationale, rollout_disposition
```

交付物合同的判定顺序是：用户可见 `instruction.md`、Judge 实际读取的文件、`tests/expectations.json`、`task.toml`、`metadata/task.json`。后三者必须在改造后与指令一致。只有五处证据对齐后，才执行目录移动或重写。

## 60 道题的处置矩阵

### 直接保留：45 道

其余 45 道都已经只要求一个 XLSX，保留在 `tasks/xlsx`。多文件输入、专业模型、图片/PDF 抽取或复杂计算不构成 Cross Office 理由。它们仍要经过输出合同审计，但不改写任务意图。

其中 `cowork-bench-insales-inventory-reorder-forecast` 的指令还提到 `supplier_lead_times.json` 和 `sales_velocity.xlsx`。当前 metadata 只把 `Inventory_Forecast.xlsx` 列为输出；审计时必须把前两者明确写成 solver 不需提交的中间输入/计算参考，或在三份配置中一致地列为交付。未完成这一步前，不能把它计入“已对齐的纯 XLSX”数量。

### 保留在 Excel，改为一个 XLSX：7 道

| 任务 | 当前多余输出 | 改造后的单一交付 | 保留的核心工作流 |
|---|---|---|---|
| `officebench-1-11-score-split` | `midterm1.xlsx`、`midterm2.xlsx` | `midterm_split.xlsx`，含两个成绩工作表 | 拆分两次考试、保留姓名和行序。 |
| `cowork-bench-json-data-cleaning-pipeline` | Python、四份 JSON、两个 XLSX | 一份数据清洗工作簿，含清洗结果、分析、库存状态、组合视图、看板和运行摘要 | 从固定 JSON 去重、规范化、聚合、低库存判断和可追溯分析。 |
| `cowork-bench-sf-regional-sales-forecast` | `Forecast_Report.docx` | `Sales_Forecast.xlsx`，增加报告/说明工作表 | 五区域预测模型、方法、假设、区域解读和建议都由工作簿承载。 |
| `cowork-bench-sf-sla-breach-audit` | 邮件 TXT | `SLA_Breach_Audit.xlsx`，增加可编辑的升级草稿工作表 | SLA 连接、违约聚合、公式与升级事实。 |
| `forte-finance-018-receivables-payables-consolidation` | Python、两份 CSV、TXT | 一份跨期往来整合工作簿 | 未付/未收、三期僵尸账核查、来源追踪和动态公式。 |
| `officebench-3-8-revenue-filter-report` | `report.docx` | 一份收入筛选工作簿，含筛选结果与分析摘要工作表 | 阈值筛选、删除行和统计说明。 |
| `spreadsheetbench2-09-01-lbo-formula-repair` | 修复器 Python、JSON 日志 | `09_01_repaired.xlsx` | 原题的模型审计和公式修复；变异/replay 从 solver 交付移到评测侧。 |

这七题必须按单题制作规范重新制作 Golden、Rubric、Judge 和负例；尤其不能只删除输出名。现有 rollout 对这些重写后的交付物无效，必须以新任务版本重新 rollout。

### Cross Office 交接：5 道（本计划不执行迁移）

| 任务 | 交接理由 | 建议家族 |
|---|---|---|
| `cowork-bench-customer-satisfaction-analysis-report` | 两份 Excel 与两份 Word 的指标、行动和叙述必须一致。 | `shared_source_multi_output` |
| `cowork-bench-faculty-workload-review` | 两份 Excel 与一份 PPT 共同完成教学负荷复核。 | `extract_analyze_present` |
| `cowork-bench-financial-audit-reconciliation` | XLSX、DOCX、EML、ICS 共同表达同一审计结论。 | `cross_file_consistency` |
| `cowork-bench-supplier-contract-compliance` | 两份 Excel、两份 Word、邮件和日历共同组成合同审计包。 | `template_driven_package` |
| `officebench-3-52-agenda-calendar-email` | 日历事件与三封邮件是相互依赖的异质交付。 | `shared_source_multi_output` |

这 5 道由学长负责迁移。本轮只在审计表中记录 `handoff_to_senior`，不移动目录、不改 instruction、Golden、Rubric、Judge 或元数据，也不把它们从当前 Excel 活动集移除；历史 rollout 继续保留，待学长完成 Cross Office 重分类后再由其生成新的领域 manifest。

### 迁移到 Word：1 道

| 任务 | 迁移理由 | Word 家族建议 |
|---|---|---|
| `officebench-2-19-student-teaching-report` | 唯一交付是 `teaching.docx`，XLSX 只是输入。 | `material_to_report` |

该题与目前 Word 题库的“多格式输入、单一 DOCX 输出”规则一致。本轮不执行迁移，保留原目录和历史 rollout，待学长或 Word 负责人另行处理。

### 归档并替换：2 道

| 任务 | 退出原因 | 替换原则 |
|---|---|---|
| `officebench-1-13-graduation-check` | 唯一输出是 `answer.txt`，不是 Excel 或多交付物工作流。 | 用一个有来源、单一 XLSX 的 L2 任务替换。 |
| `officebench-2-14-score-calendar` | 唯一产物是三份同类 ICS；既不是 XLSX，也没有异质交付物。 | 用一个有来源、单一 XLSX 的 L3 任务替换。 |

归档仅把任务移入 `datasets/xlsx-archive/` 并保留其 provenance、Golden、Judge 和历史 rollout；不删除历史证据，也不把其评分混入新版 Excel 60 题统计。

## 难度配额

当前目录工作集 60 包为 L1/L2/L3/L4 = `15/17/16/12`。其中 54 道单一 XLSX 包的分布为 `15/17/13/9`；5 道 Cross Office 和 1 道 Word 包保留其原标签，等待学长迁移时再由其负责所在领域的配额。我们只归档 2 道非 XLSX 输出题，恢复 2 道已审计的单一 XLSX 候选；目录工作集保持 60 包，每档均在约 15 的范围内。

恢复题是 `officebench-2-15-calendar-to-xlsx`（L3）和 `sb2-template-14-09-concession-revenue`（L2）。这只是作者结构假设，未用 rollout 分数反推标签；最终难度仍需独立 rollout 和人工复核。

## 新候选的来源策略

不从任务家族名或业务故事编造题目。每个替换名额必须先进入候选清单，再按 `author-xlsx-harbor-tasks` 的来源审计和 Harbor 制题流程制作。

1. 本轮恢复已完成来源、输入、真值和泄漏审计的 `XLSX-C015` 与 `XLSX-C048`，不再新增未经审计的题。
2. 如果后续 Cross Office/Word 迁出造成新的数量缺口，不能为了凑数把 Cowork-Bench 的多交付包缩成一个虚假的 XLSX；应从有明确再分发权利的 SpreadsheetBench-2 或其他公开、单工作簿来源重新筛选。
3. `Workspace-Bench-Lite` 和 BlueFin 现有 L4 候选仍为 conditional：前者缺资产许可，后者是 CC-BY-NC。未解除阻断前不能制题或计入配额。
4. 候选真正进入 `tasks/xlsx` 前，必须有一个原生 XLSX 作为唯一 solver 输出；处理脚本、审计日志和变异工具只能留在参考解法或 Judge 侧，除非任务本质确为 Cross Office。

## 执行顺序与门槛

1. 生成并人工复核输出合同审计表，修正 `instruction.md`、`task.toml`、`metadata/task.json` 和 Judge 输出约束的不一致。
2. 对 7 个 Excel 保留题逐题重新制题：每题均重做 evidence-to-requirement map、Golden、至少 6 条原子 Rubric、独立 Judge、空/损坏/未改输入/中等缺陷和必要的重算或渲染检查。
3. 对 2 个退出题执行可恢复归档；只在替换题完成基础验证后才从新版 Excel 60 的活动清单移除，避免数量断档。
4. 恢复 2 个已审计候选，并保留 `pending_human_review`；不新增未经资产审计的替代题。
5. 更新 Excel 活动集、候选映射和领域统计；Cross Office 交接题和历史 rollout manifest 原样保留。
6. 对改造和新制的题运行新的、隔离的 rollout；旧版交付物不同的结果不得沿用。随后按新的活动 60 题生成新 manifest 和统计报告。

数量核对：原 60 道中 7 道改造成单一 XLSX；5 道 Cross Office 和 1 道 Word 题本轮保持原目录；2 道非 XLSX 输出题归档；恢复 2 道已审计单一 XLSX 候选后，`tasks/xlsx` 仍为 60 道。若学长后续移动 5 道 Cross Office 题，需由其负责重新补题和生成新的领域 manifest。

## 完成定义

本轮 Excel 重整完成的条件是：当前 54 道 Excel 包均只要求一个 XLSX；5 道 Cross Office 和 1 道 Word 包只留下交接记录、不在本轮改造或计入 Excel 合同完成数；目录工作集的难度标签为每档约 15 道；所有改造、归档和恢复都有可追溯审计记录；题包通过资产、Golden、负例和渲染/重算验证；改造题与恢复题保留 `pending_human_review`。本轮未对重制或恢复题声称已有可复用的新模型 rollout。
