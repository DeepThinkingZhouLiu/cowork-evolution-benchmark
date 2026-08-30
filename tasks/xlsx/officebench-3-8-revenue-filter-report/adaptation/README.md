# C020 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=advanced` 是 L3 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/3-8/subtasks/0.json`。
- 原始 Query：`read company revenues, delete years that revenues are lower than $2000000 and save in new.xlsx, generate reports report.docx`。
- 输入：`company_revenues.xlsx`，SHA256 `b7a5d913221fafb9f545e9367df2bcf07e665b93baefe726788a88532c937e28`，5,254 bytes。
- 许可：OfficeBench Apache-2.0；上游 LICENSE 和完整任务 JSON 保留在本目录。

## 独立真值和改造边界

输入含 2004--2023 共 20 行。严格按 `< 2,000,000` 删除 2006=`1,193,225`、2009=`1,791,112`、2017=`1,463,500`，保留 17 行且顺序不变。上游 evaluator 只抽查输出第 4 行收入 `2,793,265` 和 Word 包含 `revenues`；本地 Judge 从输入重算完整集合。

改编保留两个原始输出名，只补固定输出目录、等于阈值时保留、原始顺序/数值保持和报告最低事实范围。没有要求特定报告措辞、配色或上游 Golden 布局。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 按 2,000,000 阈值保留完整集合 | 原 Query 和输入 20 行 | R002/R003 |
| 报告说明阈值与行数 | 原 Query 的 report.docx 交付要求；输入重算 | R004 |
| 报告列出删除年份和收入 | 输入中三条被删除记录 | R005 |
| 两个交付物一致 | 同一输入和同一筛选规则 | R006 |
| 输入来源固定 | 上游文件 SHA256 | R008 |

## 验证记录

已执行 `validate-task`、`validate-task-assets` 和全库 `validate-public-tasks tasks`，均通过；全库当前为 65 个公开端到端任务。Golden 与从空目录运行 `solution/solve.sh` 的 fresh replay 均由标准 `tests/test.sh` 判为 `1.000000`，同一 Golden 连续 5 次均为 `1.000000`。

负例结果：空输出/损坏 DOCX/只缺报告=`0.000000`，未筛选输入=`0.450000`，遗漏一个应保留年份=`0.500000`，报告把保留行数写成 16 而 Excel 仍为 17=`0.750000`；全部失败。Judge 从输入重算完整 17 行集合，并按 Word 表格行核对三个删除年份及其收入，不要求逐字匹配。

LibreOffice 7.3.7.2 已分别把 Excel 和 Word 打开并导出 1 页 PDF。Excel 显示 17 行且首尾年份为 2004/2023；Word 的阈值、20/17/3 摘要和三行删除表完整可见，人工预览未发现截断或重叠。

## 难度与人工决策

- 单一输入但需要 Excel 筛选、Word 报告和跨文件一致性，作者暂定 L3；如果 rollout 通过率明显高于 L3 范围，应降为 L2。
- leakage group 为 `officebench-revenue-report`，同收入工作簿及其任务变体不得跨切分。
- 报告最低事实范围、最终难度、许可结论和是否接受均待人工复核；无辅助 rollout 尚未得到模型分数。
