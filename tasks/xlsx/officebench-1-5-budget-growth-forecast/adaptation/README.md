# C007 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-5/subtasks/3.json`。
- 原始 Query：`if budget grow 10% per year, add 'predicted budget' for 3 years later to the last row in budget excel file`。
- 输入：`company_budget.xlsx`，本包改名为 `company_budget_input.xlsx`；SHA256 `7b14e19c56e1b5185f61e16cda1f1b00d424d43cb41803ba5d7b9141a4a160a3`，8,509 bytes。
- 许可：OfficeBench Apache-2.0；上游 LICENSE 和完整任务 JSON 保留在本目录。

## 独立真值和改造边界

输入的三项预算为 2,000,000、50,000、1,000，合计 2,051,000。按年增长 10% 三年复利计算为 `2,051,000 * 1.10^3 = 2,729,881`，与上游 evaluator 的目标值一致。Harbor 改编只明确输入名、输出路径、复利口径、原生公式和非目标保持；没有增加业务数据。公式要求是可重算电子表格的审计层，需人工确认其强度是否合适。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 最后一行标签 | 原 Query 和上游 evaluator | R002 |
| 三年复利值 | 输入三项金额独立合计；evaluator 交叉核验 | R003 |
| 公式引用输入且可重算 | 用户要求增长计算；原生 Excel 交付 | R005 |
| 原始区域保持 | 输入工作簿逐格比较 | R001/R004 |
| 输入来源固定 | 上游文件 SHA256 | R008 |

## 验证记录

已执行 `validate-task`、`validate-task-assets` 和全库 `validate-public-tasks tasks`，均通过；全库当前为 65 个公开端到端任务。`solution/solve.sh` 从空目录生成的 fresh replay 和仓库 Golden 均由 `tests/test.sh` 判为 `1.000000`，同一 Golden 连续 5 次均为 `1.000000`。

负例结果：空输出/损坏 XLSX=`0.000000`，未修改输入=`0.250000`，错误两年复利公式=`0.600000`，硬编码正确数值=`0.850000` 但 R005 门槛失败，篡改原预算标签=`0.850000` 但 R004 门槛失败。Judge 从输入重算总额，并把首项预算增加 1,000 后通过 LibreOffice 重算验证公式行为，不读取 Golden。

LibreOffice 7.3.7.2 已打开并导出 1 页 PDF；`predicted budget` 和 `2,729,881` 可见，无截断。人工预览确认源表的简洁样式与新行一致。

## 难度与人工决策

- 单工作簿、一个追加区域，但包含复利公式、保持约束和变异重算，作者暂定 L2；若 rollout 显示通过率过高，应降为 L1。
- leakage group 为 `officebench-budget-formula`，OfficeBench 同预算工作簿变体不得跨切分。
- 公式审计层、最终难度、许可结论和是否接受均待人工复核；无辅助 rollout 尚未得到模型分数。
