# Course Enrollment Analytics Dashboard (XLSX-only)

请基于输入目录中的 Canvas 课程、选课、作业和提交快照，生成一个可重算的 Excel 工作簿。输入是同一 Cowork-Bench 固定快照的离线投影；不得使用网络、外部服务或凭空补充课程/成绩。

只交付 `/app/output/Course_Enrollment_Analytics.xlsx`，输出目录不得包含其他文件。工作簿按以下顺序包含五张原生可编辑工作表：

1. `Course_Metrics`：每门有至少一条非空成绩提交的课程一行，列为 `Course_ID`、`Course_Name`、`Course_Code`、`Total_Enrollment`、`Student_Enrollment`、`Graded_Students`、`Submission_Count`、`Graded_Submissions`、`Avg_Score`、`Completion_Rate`、`Late_Rate`、`Risk_Level`。`Total_Enrollment` 为该课程所有选课行数；`Student_Enrollment` 只计 `type` 含 Student 的行；`Graded_Students` 为至少有一条非空成绩的不同学生数；`Submission_Count` 为该课程所有提交行；`Graded_Submissions` 为 `workflow_state=graded` 的提交数。平均分为选课记录 `current_score` 的平均，完成率为已评分提交数除以全部提交数，迟交率为 `late=true` 的提交数除以全部提交数；百分比保留 1 位小数。按 `Avg_Score` 升序排列。

2. `Risk_Distribution`：列为 `Risk_Level`、`Course_Count`、`Student_Enrollment`。按 `Critical`、`At Risk`、`On Track` 排列，课程风险按平均分 `<60`、`60–75`（含 60、不含 75）、`>=75` 判定。统计必须使用原生公式引用 `Course_Metrics`，而不是静态硬编码。

3. `Submission_Quality`：每门课程一行，列为 `Course_Name`、`Submission_Count`、`Graded_Submissions`、`Completion_Rate`、`Late_Submissions`、`Late_Rate`；指标应与 `Course_Metrics` 一致，并使用公式链接或等价的可编辑计算。

4. `Risk_Action_Plan`：列为 `Risk_Level`、`Action`、`Timeline`、`Responsible`，按 `Critical`、`At Risk`、`On Track` 排列。`Critical` 使用“一对一学业辅导/一周内/Academic Advisor”；`At Risk` 使用“小组辅导/两周内/Course Tutor”；`On Track` 使用“自定进度学习资源/一个月内/Student Success Office”。这些行动来自 `scoring_model.json` 的干预规则，可保留英文原文并补充中文说明。

5. `Source_Notes`：记录输入文件名、来源角色、风险阈值、计算口径和快照说明，便于复核；不要写入未由输入支持的姓名、原因、趋势或效果。

工作簿应启用自动完整重算，使用原生 Excel 表格、冻结首行、加粗表头、合理列宽和可读数字格式。公式必须在 LibreOffice 中可重算。不要把截图或只读 PDF 当作表格替代品。
