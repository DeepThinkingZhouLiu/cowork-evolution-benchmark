# C085 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：TheAgentCompany commit `98b68ef82a47690c316f42fddb05baafaab56851`，`workspaces/tasks/finance-budget-variance`。
- 输入：`budget.xlsx`（SHA256 `c9578fd151590966ee1003559132fcf81ede7aa249393ec8d9e3a57ce063607d`，12,945 bytes）和 `actual_spending.xlsx`（SHA256 `060d0ecadab3863cc71fcb838e10e84cca73c81beb4d407b063ee410f5e2d139`，14,466 bytes）。
- 许可：MIT；上游 task、checkpoints、evaluator 和 LICENSE 均保留在本目录。

## 独立真值和改造边界

两份输入各含 420 个完全对齐的 `Department/Category/Month` 键，预算无零值。按 `Actual - Budget` 和差异百分比重算，并严格使用 `abs(percentage)>10 AND abs(amount)>5000`，得到 8 条重大差异：

| 标识 | 差异金额 | 差异百分比 |
|---|---:|---:|
| Engineering_Salaries_202304 | 36,364.08 | 23.72% |
| Engineering_Salaries_202306 | 20,861.66 | 12.37% |
| Engineering_Salaries_202307 | 34,644.93 | 20.55% |
| Engineering_Salaries_202312 | -25,678.08 | -13.96% |
| Engineering_Software_202302 | 5,077.12 | 28.11% |
| Marketing_Salaries_202304 | -28,672.62 | -22.36% |
| HR_Salaries_202308 | -32,507.49 | -22.84% |
| Operations_Salaries_202306 | -19,680.57 | -19.91% |

上游 evaluator 只要求 `Flagged_Variances` 标识列。本地包保留该列，并添加用户原请求已要求计算的预算、实际、金额差异和百分比审计列。月度预算数值本身已经随月份变化，因此“调整季节性”的有证据解释是使用同月预算，不另造季节系数。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 8 个重大差异标识 | 两份输入独立联结和双阈值规则 | R002/R005 |
| 部门、类别、月份和源金额 | 两份输入同键记录 | R003 |
| 差异金额与百分比 | 原 task 公式语义和输入重算 | R004 |
| 标识格式与顺序 | 原 task/evaluator 和预算输入顺序 | R006 |
| 输入来源固定 | 两个上游文件 SHA256 | R008 |

## 验证记录

已执行 `validate-task`、`validate-task-assets` 和全库 `validate-public-tasks tasks`，均通过；全库当前为 65 个公开端到端任务。Golden 与最终 `solution/solve.sh` fresh replay 均由标准 `tests/test.sh` 判为 `1.000000`，同一 Golden 连续 5 次均为 `1.000000`。

负例结果：空输出/损坏 XLSX/把 budget.xlsx 直接改名输出=`0.000000`；遗漏一条重大差异失败；把一条来源实际金额改错=`0.850000` 且 R003 门槛失败；把差异金额改错=`0.800000` 且 R004 门槛失败；增加输入中 `Operations_Software_202308` 作为误标行=`0.250000`，该行百分比为 -28.87% 但金额仅 -3,954.70，未同时超过双阈值。Judge 每次从 420 个联结键重算，不读取 Golden 或上游期望列表。

LibreOffice 7.3.7.2 已打开工作簿并导出 1 页横向 PDF。8 行、8 列和全部正负金额/百分比可见；人工预览确认无横向分页、截断或重叠。

## 难度与人工决策

- 两个 420 行工作簿、精确联结、双阈值和符号处理构成单应用多阶段任务，作者暂定 L2。
- leakage group 为 `tac-budget-variance`；同一预算/实际 fixture 的变体不得跨切分。
- 季节性解释、输出审计列、最终难度、许可结论和是否接受均待人工复核；无辅助 rollout 尚未得到模型分数。
