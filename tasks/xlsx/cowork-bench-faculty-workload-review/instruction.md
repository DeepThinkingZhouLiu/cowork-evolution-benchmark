# 教师教学负荷分析与复核

你是教务分析师。请使用 `/app/data/input_files` 中的固定快照完成一次可复核的教学负荷分析：

- `canvas_courses.csv` 是课程快照，包含课程代码、学科领域以及课程的学生数。
- `canvas_teacher_enrollments.csv` 是教师与课程的授课关系。
- `canvas_student_enrollments.csv` 是学生选课记录；只统计这些记录中的学生选课行。
- `canvas_assignments.csv` 是课程作业记录；按课程统计作业数。
- `workload_standards.json` 和 `workload_policy.pdf` 给出课程数、批改时间、学期周数和超负荷规则。
- `faculty_directory.csv` 是教师目录，用于核对机构背景；不要改写输入文件。

对每位在教师授课关系中出现的教师，统计不同课程数、这些课程的学生选课总数和作业总数。估算学期批改小时数为 `Total_Students * grading_hours_per_student`，每周批改小时数为学期小时数除以 `semester_weeks`。当课程数大于 `max_courses_per_semester` 或每周批改小时数大于 `max_weekly_hours` 时，将 `Overloaded_YN` 设为 `Yes`，否则设为 `No`。按授课关系中的 `Login_ID` 字母顺序输出；`Instructor` 单元格保留 Canvas 快照中的显示姓名。

在 `/app/output/Faculty_Workload.xlsx` 中创建两个工作表：

1. `Instructor Load`：列为 `Instructor`、`Courses_Count`、`Total_Students`、`Total_Assignments`、`Est_Grading_Hours`、`Weekly_Grading_Hours`、`Overloaded_YN`。
2. `Department Summary`：按课程代码的学科领域聚合授课关系，列为 `Department`、`Instructor_Count`、`Course_Count`、`Total_Students`。每个教师-课程授课关系计入该领域的学生负荷；因此同一课程由多位教师授课时，其学生数会分别计入各授课关系，但 `Course_Count` 只计不同课程代码。

将 `Instructor Load` 数据另存为 `/app/output/Faculty_Workload_Analysis.xlsx`，作为题目要求的本地可编辑共享分析表替代品，工作簿标题设为 `Faculty Workload Analysis`。

另建 `/app/output/Workload_Review.pptx`，至少包含：标题页、教师总数和超负荷教师数、按学科领域的汇总，以及针对超负荷领域的课程/批改量再分配建议。演示文稿中的数字必须与两个 Excel 工作簿一致。Excel 表格和 PPT 文本应保持原生可编辑，不要用截图替代。
