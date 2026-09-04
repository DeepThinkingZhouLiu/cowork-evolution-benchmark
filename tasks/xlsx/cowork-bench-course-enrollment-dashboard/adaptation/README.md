# Cowork-Bench course-enrollment-analytics-dashboard 逐题改造记录

## 来源与许可

- 上游仓库：`0717376/cowork_bench`
- 固定 revision：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 原任务：`tasks/finalpool/course-enrollment-analytics-dashboard`
- 源 SQL：`db/init.sql.gz`，SHA256 `bd9c45b5ec9072b9f301b7df5b3e9ac65540ffcd53485329ebea44b0c3710da4`
- 许可：仓库 Apache-2.0；副本保存在 `LICENSE-Cowork-Bench`。

原任务要求从 Canvas 动态读取课程、选课与成绩，并写入 Google Sheet、Teamly、邮件和日历。本题采用 offline workflow reconstruction：从同一不可变 SQL 快照投影课程 16/17 的任务相关字段，把外部副作用和四类报告压缩到一个本地可编辑 XLSX。不会声称创建外部页面、邮件或日历事件。

## 输入与证据 ledger

| evidence_id | 上游路径 / 本地路径 | 角色 | 处理 | SHA256 |
|---|---|---|---|---|
| E1 | `db/init.sql.gz` → `data/input_files/canvas_courses.csv` | 课程身份、代码、状态 | `extract_canvas_snapshot.py` 投影 16/17 | `72f17c5402df529b0d8c822fe87050c4c3c75c5f7c220d963fb592d957557162` |
| E2 | `db/init.sql.gz` → `canvas_enrollments.csv` | 全部选课行、Student 类型和 current_score | 仅保留 16/17；未新增记录 | `afb3b26cb6a101d741c3cf3560b03ccc2f17047b79fc70b01b7864f9957cad3e` |
| E3 | `db/init.sql.gz` → `canvas_assignments.csv` | 作业到课程的映射 | 保留 26 个作业 | `98573aae81792ae965503241dff2a162a9170d20d86daed8a337c72e765d6d16` |
| E4 | `db/init.sql.gz` → `canvas_submissions.csv` | 提交、workflow_state、late 和 score | 保留上述作业的 32,424 行 | `11f34a9c4fe74934b1f8c9576358f94d45a7ae63f60efa66f02365ec0b6c2125` |
| E5 | 原任务 `scoring_model.json` / `Retention_Policy.pdf` | 风险阈值、行动和口径 | 原样保留/结构化转录 | `ef019f895d416c3495c59d43e867125aea63b7f4fda929e2959771a980483c9e` / `1e851c80117c065d3535716b98b4486e41e3b6a7a7372662947a77f501baa4d8` |

## 改造与真值

- 保留原题的两个有成绩课程、按课程/学生平均 `current_score` 的风险规则、提交完成率与迟交率。
- 课程 16：总选课 2,285，学生选课 2,283，已评分学生 1,872，提交 16,240，已评分提交 16,228，平均分 74.94，完成率 99.9%，迟交率 10.0%，`At Risk`。
- 课程 17：总选课 2,369，学生选课 2,365，已评分学生 1,842，提交 16,184，已评分提交 16,175，平均分 76.51，完成率 99.9%，迟交率 10.3%，`On Track`。
- 上游任务的动态评测和旧 Golden 与当前源行存在 enrollment 计数差异；本题以输入行重算为准，不复制过期 Golden 数字。
- 需要至少有一条成绩的课程才进入指标表；这是原 evaluator 对无成绩课程的处理。

## 交付与 Judge

交付 `Course_Enrollment_Analytics.xlsx`，包含 Course_Metrics、Risk_Distribution、Submission_Quality、Risk_Action_Plan 和 Source_Notes 五张表。Judge 从四个 Canvas CSV 独立重算数值，检查风险边界、公式、表格结构、输入哈希，并尝试 LibreOffice 重算。负例包括缺失/损坏输出、空表、额外文件、静态风险分布值和未保留行动规则。

## 难度与未决事项

作者假设：L4/expert；依据是四个异构源文件、动态状态投影、课程/学生/提交多级汇总、外部副作用压缩、公式与跨表一致性。状态为 `uncalibrated`，评审状态为 `pending_human_review`。未执行模型 rollout；评审者仍需确认 XLSX-only 投影是否保留了原任务期望的外部协作语义，以及是否接受只覆盖原题有成绩的 16/17 课程。
