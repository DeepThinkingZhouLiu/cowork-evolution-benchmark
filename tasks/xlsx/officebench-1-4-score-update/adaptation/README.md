# C001 逐题改造记录

状态：`pending_human_review`。这是 pilot 的 L1 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench，commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，任务 `tasks/1-4/subtasks/0.json`。
- 原始 Query：`change Bob's midterm1 score 100 to score.xlsx excel file`，原 evaluator 检查 `Sheet1!B12=100`。
- 输入：上游 `score.xlsx` 的字节副本，Harbor 中命名为 `score_input.xlsx` 以避免与输出 `score.xlsx` 触发同名 Golden 检查；SHA256 `a5c8026dc09f38411c6b8783dad74f7f66f08387eeb21b9d8eb7f28a90dd5b6e`，5,334 bytes。
- 许可：上游仓库 Apache-2.0；来源 LICENSE 和原任务记录由作者侧保留。

## 改造边界

保留了单工作簿单值修改意图，只补充固定输入名、`/app/output/score.xlsx`、原生 XLSX 和非目标保持要求。输入事实显示 Bob 在第 12 行、`midterm1=97`；Golden 只把 `B12` 改为 100。没有从 Golden 推导其他事实。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| Bob 的 midterm1 为 100 | 原 Query、输入第 12 行、上游 evaluator | R002 |
| 其他单元格不变 | 输入工作簿逐格比较 | R003 |
| 工作簿可编辑且结构/样式保持 | 输入 OOXML 与输出 workbook 检查 | R001/R004/R006 |

## 验证记录

`solution/solve.sh` 从输入生成 Golden；`tests/judge.py` 独立读取输入重算目标并逐格检查非目标区域。已执行：

- `validate-task`、`validate-task-assets`：通过；输入副本改名为 `score_input.xlsx` 后无 Golden 泄漏。
- Golden 与 fresh replay：Judge `1.000000`。
- 空输出 `0.000000`、损坏 XLSX `0.000000`、未修改输入 `0.833333`（R002 门槛失败）。
- LibreOffice 7.3.7.2 独立 profile 打开并导出 PDF，文本可提取且包含 Bob/100。

## 风险与人工决策

- 上游 evaluator 只检查一个单元格；本题新增的非目标保持检查是通用文件完整性约束，需要人工确认是否过严。
- OfficeBench 其他同 workbook 成绩题属于 leakage group `officebench-score-basic`，不得跨切分复用。
- Codex/OpenCode 难度 rollout 尚未执行；`difficulty_tier=basic` 仍为 `uncalibrated` 假设。
