# C005 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=basic` 是 L1 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-11/subtasks/0.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-11/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `b0b81274ef8087f31e6920d8ff7ae3358c082979533f71b2fde134075744c973`，4 条学生记录。
- 上游 `reference/sorted_score.xlsx` SHA256 `481fa316f5fd1916727fc98209e9104ded4d9d76ca9433f6a86d49efe8a9e906` 已保留作来源审计证据；Judge 不比较其 ZIP 字节。

## 独立真值和改造边界

从输入读取 `Name`、`midterm1`、`midterm2` 三列，将数据行按第二列数值降序稳定排序；输入顺序得到 Alice(78)、Liam(74)、Jack(71)、Ivy(64)。改造明确表头、列顺序、结构和格式保持，不额外要求 Excel 排序按钮或特定内部 XML。Judge 重新排序输入并比较完整值矩阵。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| midterm1 降序 | 原 Query、输入数值 | R002/R003 |
| 所有学生和成绩保留 | 输入工作簿 | R001/R003 |
| 样式与可编辑交付 | 输入/任务语义 | R004/R006 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward 均为 `1.000000`。空输出/损坏 XLSX=`0.000000`，未排序输入=`0.600000`（R002 门槛失败）；Judge 独立重排输入并拒绝按 midterm2 排序或遗漏行。LibreOffice 7.3.7.2 成功打开并导出 1 页 PDF，Alice 排在首行。

## 难度与人工决策

单工作簿四行排序并另存，作者暂定 L1；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-sort`。
