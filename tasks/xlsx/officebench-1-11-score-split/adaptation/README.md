# C011 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。版本 `0.2.0` 将原两个同类型交付合并为一个工作簿的两个工作表。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-11/subtasks/1.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-11/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973`。
- 上游参考 `midterm1.xlsx`、`midterm2.xlsx` 已保留在 `adaptation/`，哈希分别为 `3f4ead4dcffa430920fd445a196f1ec324b65f07c71b0813f3041fa4117f4568` 和 `1cf70058ee0716cf25942fbf29533e07ad85c7d8d56d634d58f80a35d0f5a80d`。

## 独立真值和改造边界

输入的三列被投影为一个工作簿中的两个两列表格：Name+midterm1、Name+midterm2，分别位于 `midterm1` 和 `midterm2` 工作表。Judge 从输入列独立生成期望矩阵，检查唯一输出的工作表顺序、列标题、值、结构、样式和可编辑性，不读取参考 Golden。原双工作簿参考文件继续保留在 `adaptation/` 作为上游适配证据；新版 Golden 由参考解法重新生成。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 一个目标工作簿和两个工作表 | 原 Query、输入表头、同类型合并规则 | R001/R002/R003 |
| 对应成绩值和顺序 | 输入工作簿独立投影 | R002/R003/R004 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

已运行 `validate-task`、`validate-task-assets`、Golden/fresh replay，以及空/损坏/未修改/交换工作表负例。结构校验与资产校验通过；Golden 和 fresh replay 的 Judge reward 均为 `1.000000`，新 Judge 不再接受旧的两个文件合同。LibreOffice 检查确认一个工作簿包含两个可见、可编辑工作表。

## 难度与人工决策

单输入到两个原生工作簿的结构拆分，作者暂定 L2；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-split`。
