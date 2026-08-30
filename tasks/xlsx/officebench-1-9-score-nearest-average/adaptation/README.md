# C010 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-9/subtasks/1.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-9/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `12fce0b4e70006d9dc40745b532d75ee2e3c0ffc638cca2393bf2a9d5770226d`，20 条学生记录。
- 输入 `midterm1` 总和为 1601，算术平均值为 80.05；与平均值绝对差最小的是输入第 4 行 `Alice,78,75`。

## 独立真值和改造边界

原 Query 要求找出最接近平均值的学生并复制整行。改造明确平均值覆盖所有数据行、并列时保留输入中最先出现者，再追加一行；这些规则可从输入和 Query 直接核验，没有使用上游 80.5 的冲突值。Judge 从输入重算平均值、最近行和原区域保持，不读取 Golden。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 平均值和最近行 | 原 Query、输入 20 行 | R002/R003 |
| 追加完整复制行 | 原 Query、输入最近记录 | R003 |
| 原数据和样式保持 | 输入工作簿 | R001/R004 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward `1.000000`。空输出/损坏 XLSX=`0.000000`，未修改输入=`0.100000`，追加错误学生行=`0.450000`；Judge 拒绝非 22 行追加、错误最近行和原区篡改。LibreOffice 7.3.7.2 成功打开并导出 1 页 PDF，追加 Alice 记录可见。

## 难度与人工决策

单工作簿平均值、绝对差和行复制，作者暂定 L2；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-analysis`。
