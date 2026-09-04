# C006 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=basic` 是 L1 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-14/subtasks/1.json`；上游任务 JSON、reference 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-14/testbed/data/score.xlsx`，改名为 `score_input.xlsx`；SHA256 `6b98c5b9d854ab3d5fce62632a58a49c58559fff04435bb59d68fe1f0c6dd255`。
- 上游 `reference/score_swap_column.xlsx` SHA256 `331efef891b5d06b332aac77c382a9cd6887dc1c2c7cb1a91b4677f3f93c9a9e`，仅作来源审计证据。

## 独立真值和改造边界

输入表头为 `Name, midterm1, midterm2`，4 行学生数据。交换规则独立得到 `Name, midterm2, midterm1` 及逐行交换后的值。Judge 从输入重算目标矩阵和逆变换，不比较 reference ZIP 字节；改造只明确输出路径、结构保持和原生 Excel 交付。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| 两列位置和数据交换 | 原 Query、输入三列表 | R002/R003 |
| 结构、顺序和格式保持 | 输入工作簿 | R001/R004/R006 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward `1.000000`。空输出/损坏 XLSX=`0.000000`，未修改输入=`0.400000`，只交换数据但不交换表头=`0.600000`（R002 门槛失败）。LibreOffice 7.3.7.2 成功打开并导出 1 页 PDF，表头和 Liam 的交换值可见。

## 难度与人工决策

单工作簿列重排，作者暂定 L1；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-score-structure`。
