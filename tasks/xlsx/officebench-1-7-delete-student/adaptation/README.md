# C003 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=basic` 是 L1 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-7/subtasks/0.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-7/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `023d650c0ce6192318d49f53cf099fddb109eb026953c96f3372e66fcbe01bc7`，20 条学生记录。
- 输入中 Alice 的一行是 `Alice,78,75`；其余记录按原顺序保留。

## 独立真值和改造边界

删除规则由 Query 指定为删除 Alice 及其成绩。Judge 从输入过滤姓名为 Alice 的数据行，并逐行比较输出；它不把上游的 `evaluate_diff_contain_text` 关键字检查当作真值，也不读取 Golden。改造仅补充明确输出路径、结构保持和原生工作簿要求。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 删除 Alice 整行 | 原 Query、输入 Alice 行 | R002/R003 |
| 其他记录和顺序保持 | 输入工作簿 | R001/R003/R004 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward 均为 `1.000000`。空输出/损坏 XLSX=`0.000000`，未修改输入=`0.250000`，删除 Bob 而保留 Alice=`0.450000`；两类错误均触发 R002/R003 失败。LibreOffice 7.3.7.2 成功打开并导出 1 页 PDF，输出不含 Alice。

## 难度与人工决策

单工作簿单行删除，作者暂定 L1；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-basic`，与同一成绩基础工作簿变体隔离。
