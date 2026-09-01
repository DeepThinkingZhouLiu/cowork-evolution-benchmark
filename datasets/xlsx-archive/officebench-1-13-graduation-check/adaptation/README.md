# C012 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=intermediate` 是 L2 作者假设，不是最终难度或发布决定。

## 来源和证据

- 来源：OfficeBench commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`，`tasks/1-13/subtasks/2.json`；上游任务 JSON 和 Apache-2.0 LICENSE 保留在本目录。
- 输入：`tasks/1-13/testbed/data/Thomas_Robinson_Course_List.xlsx`；SHA256 `5a495d6f9d3c319207dc1b9ed002672a42eee30c3f25741b31ca4e7df399faeb`，15 条课程记录。
- 输入中 Course ID 以 `CS ` 开头的课程为 5 门，以 `MATH ` 开头的课程为 5 门。

## 独立真值和改造边界

毕业规则是至少 8 门 CS 且至少 5 门 MATH。输入满足数学门槛但不满足 CS 门槛，因此独立结论为 `No`。改造将原先仅检查包含关键字的 `answer.txt` 明确为 UTF-8 单行 `Yes`/`No`，没有添加解释文字或修改输入。Judge 从输入重新统计前缀并重算合取规则；`solution/Golden/course_source_evidence.xlsx` 只是 validator 所需的隐藏输入证据副本，不是 solver 输出。

## 证据到要求映射

| 要求 | 证据 | Judge |
|---|---|---|
| CS/MATH 课程数 | 输入课程表 | R003/R004 |
| 至少 8 且至少 5 的规则 | 原 Query | R005 |
| 单行答案文件 | 原 Query、适配边界 | R001/R002 |
| 输入固定 | 上游文件 SHA256 | R006 |

## 验证记录

`validate-task`、`validate-task-assets` 和 Golden/fresh replay 均通过，Judge reward `1.000000`。空输出/损坏文本=`0.000000`，输出 `Yes`=`0.850000`（规则失败），输出 `No because CS=5`=`0.650000`（单行格式失败）。LibreOffice 7.3.7.2 成功打开并导出输入课程表 1 页 PDF；文本 Golden 为严格 `No` 单行。

## 难度与人工决策

单工作簿前缀统计到确定性文本判断，作者暂定 L2；最终难度和是否接受均待无辅助 rollout 与人工复核。泄漏组为 `officebench-graduation-check`。
