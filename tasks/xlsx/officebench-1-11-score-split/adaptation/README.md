# C011 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-11/subtasks/1.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-11/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973`。
- 上游参考 `midterm1.xlsx`、`midterm2.xlsx` 已保留在 `adaptation/`，哈希分别为 `3f4ead4dcffa430920fd445a196f1ec324b65f07c71b0813f3041fa4117f4568` 和 `1cf70058ee0716cf25942fbf29533e07ad85c7d8d56d634d58f80a35d0f5a80d`。

## 独立真值和改造边界

输入的三列被投影为两个两列表格：Name+midterm1、Name+midterm2，行顺序和工作表名称保持。Judge 从输入列独立生成期望矩阵，检查两个输出的列标题、值、结构、样式和可编辑性，不读取参考 Golden。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 两个目标文件和列结构 | 原 Query、输入表头 | R001/R002/R003 |
| 对应成绩值和顺序 | 输入工作簿独立投影 | R002/R003/R004 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward 均为 `1.000000`。空输出/损坏输出=`0.000000`，交换两个输出文件=`0.400000`（R002/R003 失败）；缺少任一文件也触发 R001 门槛。LibreOffice 7.3.7.2 成功打开并分别导出两个 1 页 PDF，两个输出均可见对应考试列。

## 难度与人工决策

单输入到两个原生工作簿的结构拆分，作者暂定 L2；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-split`。
