# C002 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=basic` 是 L1 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-5/subtasks/1.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-5/testbed/data/shopping_list.xlsx`，改名为 `shopping_list_input.xlsx`；SHA256 `e6468e5452830bc4bd8767fb9a95c0ea95a65d876e306581437889e43c2a821a`，4 行 2 列，apple 的输入数量为 2。
- 许可：OfficeBench Apache-2.0；输入为固定公开 fixture，逐文件哈希记录在 `metadata/input-manifest.json`。

## 独立真值和改造边界

输入中 apple 位于 `Sheet1!A2:B2`，数量 2；“减半”独立得到 1。改造只明确离线输入/输出路径、原生工作簿和非目标保持，没有增加业务行或格式要求。Judge 从输入读取目标行并检查其余单元格，不读取 Golden。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| apple 数量为 1 | 原 Query、输入 B2=2 | R002 |
| 其他商品和结构保持 | 输入工作簿 | R001/R003/R004 |
| 输入固定 | 上游文件 SHA256 | R005 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward 均为 `1.000000`。空输出/损坏 XLSX=`0.000000`，未修改输入=`0.650000`（R002 门槛失败），把 cabbage 改为 1=`0.450000`（R002/R003 失败）。LibreOffice 7.3.7.2 成功打开并导出 1 页 PDF，apple=1 可见。

## 难度与人工决策

单工作簿单元格算术修改，作者暂定 L1；最终难度、输入副本许可确认和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-shopping-basic`。
