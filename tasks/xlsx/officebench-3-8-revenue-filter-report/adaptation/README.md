# C020 逐题改造记录

状态：`pending_human_review`。版本 `0.2.0` 的 `difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/3-8/subtasks/0.json`。
- 原始 Query：`read company revenues, delete years that revenues are lower than $2000000 and save in new.xlsx, generate reports report.docx`。
- 输入：`company_revenues.xlsx`，SHA256 `b7a5d913221fafb9f545e9367df2bcf07e665b93baefe726788a88532c937e28`，5,254 bytes。
- 许可：OfficeBench Apache-2.0；上游 LICENSE 和完整任务 JSON 保留在本目录。

## 独立真值和改造边界

输入含 2004--2023 共 20 行。严格按 `< 2,000,000` 删除 2006=`1,193,225`、2009=`1,791,112`、2017=`1,463,500`，保留 17 行且顺序不变。上游 evaluator 只抽查输出第 4 行收入 `2,793,265` 和 Word 包含 `revenues`；本地 Judge 从输入重算完整集合。

改编保留收入筛选意图和输入；原 `new.xlsx` 的筛选结果、原 `report.docx` 的阈值/行数说明及删除记录被整理为唯一 `Revenue_Filter.xlsx` 的 `Filtered Revenue`、`Filter Summary` 和 `Deleted Records` 工作表。没有要求特定报告措辞或上游 Golden 布局。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 按 2,000,000 阈值保留完整集合 | 原 Query 和输入 20 行 | R002/R003 |
| 报告说明阈值与行数 | 原 Query 的 report.docx 交付要求；输入重算 | R004 |
| 报告列出删除年份和收入 | 输入中三条被删除记录 | R005 |
| 两个交付物一致 | 同一输入和同一筛选规则 | R006 |
| 输入来源固定 | 上游文件 SHA256 | R008 |

## 验证记录

已运行 `validate-task`、`validate-task-assets`、Golden/fresh replay、空输出、损坏输出、未筛选输入和摘要/删除记录缺陷负例。结构校验与资产校验通过；Golden 和 fresh replay 的 Judge reward 均为 `1.000000`。新版 Judge 只读取唯一 XLSX，不再接受 DOCX。

负例结果：空输出/损坏工作簿/缺少摘要或删除记录=`0.000000`，未筛选输入=`0.450000`，遗漏一个应保留年份=`0.500000`，摘要把保留行数写成 16 而筛选表仍为 17=`0.750000`；全部失败。Judge 从输入重算完整 17 行集合，并核对 `Deleted Records` 工作表中的三个删除年份及其收入，不要求逐字匹配。

LibreOffice 7.3.7.2 已打开 `Revenue_Filter.xlsx` 并导出 3 页 PDF。`Filtered Revenue` 显示 17 行且首尾年份为 2004/2023；`Filter Summary` 的阈值与 20/17/3 摘要、`Deleted Records` 的三行删除记录完整可见，人工预览未发现截断或重叠。

## 难度与人工决策

- 单一输入、筛选变换加同一工作簿内的摘要/删除记录三个结果区，作者暂定 L2；若 rollout 通过率超出 L2 范围，再按结构和歧义复核。
- leakage group 为 `officebench-revenue-report`，同收入工作簿及其任务变体不得跨切分。
- 报告最低事实范围、最终难度、许可结论和是否接受均待人工复核；无辅助 rollout 尚未得到模型分数。
