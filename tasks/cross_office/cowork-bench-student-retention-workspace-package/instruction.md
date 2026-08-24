请以工作区中的 Canvas 课程数据快照、`scoring_model.json` 和 `Retention_Policy.pdf` 为唯一事实来源，完成一次学生留存风险分析工作区交付。请交付：

- `/app/output/risk_scorer.py`
- `/app/output/risk_analysis_results.json`
- `/app/output/Student_Risk_Analysis.xlsx`
- `/app/output/Intervention_Plan.docx`
- `/app/output/Student_Risk_Tracker.md`
- `/app/output/advisor_email.txt`

这是 Cowork-Bench `terminal-canvas-excel-word-teamly-email` 的离线静态改造版。上游原题从 Canvas 数据库读取课程 16/17，再写 Excel、Word、Teamly 页面和邮件。本题把同一 Apache-2.0 SQL 快照中的源行静态化为 4 份 CSV，并把 Teamly/邮件副作用分别改为 Markdown 页面和离线邮件草稿，不需要网络或外部账号。

输入职责：

- `canvas_courses.csv`：课程 16 和 17 的课程名、课程代码和源快照总人数。
- `canvas_enrollments.csv`：两门课的 4,648 条 StudentEnrollment 源记录。`enrollment_count` 必须实际按此文件计数，不得抄写旧 Golden。
- `canvas_assignments.csv`：两门课的 26 条作业记录，负责将 submission 关联到课程。
- `canvas_submissions.csv`：32,424 条提交记录；空 `score` 表示未评分，必须排除。
- `scoring_model.json` 和 `Retention_Policy.pdf`：风险阈值、干预动作、时限、责任人和报告要求。

计算口径：

- 对每个 `course_id + user_id` 计算所有非空 `score` 的算术平均；没有任何已评分提交的学生不进入风险分布分母。
- `High`：学生平均分 `< 60`；`Medium`：`60 <= 平均分 < 75`；`Low`：`平均分 >= 75`。
- `Course_Overview.avg_score` 是该课程所有非空 submission score 的平均，四舍五入保留 2 位。
- `pass_rate` 是该课程学生平均分 `>=60` 的人数除以至少有一条已评分提交的学生人数，乘 100 并保留 2 位。
- `Risk_Distribution.pct` 以两课所有“至少有一条已评分提交”的学生数为分母，保留 2 位。
- `Student_Risk_Tracker.md` 中 `Student Count` 使用已评分学生数，`Risk Level` 使用该课人数最多的风险级别；如并列则优先级为 `High > Medium > Low`。

`risk_scorer.py` 必须真实读取 CSV/JSON 输入，并支持：

```bash
python risk_scorer.py \
  --courses /app/data/input_files/canvas_courses.csv \
  --enrollments /app/data/input_files/canvas_enrollments.csv \
  --assignments /app/data/input_files/canvas_assignments.csv \
  --submissions /app/data/input_files/canvas_submissions.csv \
  --model /app/data/input_files/scoring_model.json \
  --output /app/output/risk_analysis_results.json
```

Judge 会再用一份改动某位 High 学生全部分数的 submission 快照回放该脚本，所以不得把课程统计、风险人数、百分比或文本总结硬编码为固定答案。JSON 必须包含 `course_overview`、`risk_distribution`、`at_risk_students`、`intervention_plan` 和 `student_risk_tracker`。

`Student_Risk_Analysis.xlsx` 必须包含四张原生可编辑工作表：

- `Course_Overview`：`course_id, course_name, enrollment_count, avg_score, pass_rate`，课程 16/17 各一行。
- `Risk_Distribution`：`risk_level, student_count, pct`，按 High、Medium、Low 排列；`pct` 必须用 Excel 公式从三行 `student_count` 计算。
- `At_Risk_Students`：`course_name, high_risk_count, medium_risk_count, low_risk_count`，两课各一行。
- `Intervention_Plan`：`risk_level, action, timeline, responsible`，按 High、Medium、Low 排列，必须保留 `scoring_model.json` 中的英文原文。

`Intervention_Plan.docx` 使用英文标题 `Student Retention Intervention Plan`，用俄文编写 `Executive Summary`、`Course Overview`、`Risk Distribution`、`Recommended Interventions` 和 `Method & Scope`，并与 JSON/Excel 数字一致。不得伪造学生姓名、原因、辅导成效或数据未提供的趋势。

`Student_Risk_Tracker.md` 必须以 `Student Risk Tracker` 为 H1，并包含一张两行 Markdown 表格，表头恰为 `Student Course, Risk Level, Student Count, Average Score, Pass Rate`。

`advisor_email.txt` 是离线邮件草稿，必须以下列字段开头，且正文说明 High 总人数和 High/Medium/Low 三级干预的下一步：

```text
To: academic_advisors@university.edu
Subject: Student Retention Risk Analysis - Action Required
```
