# C095 逐题改造记录

状态：`pending_human_review`。`difficulty_tier=advanced` 是 L3 作者假设，不是最终难度或发布决定；`calibration_status=uncalibrated`。

## 来源与许可

- 来源仓库：Cowork-Bench，固定 commit `d943e75bc0fc8e3b27141979300cd8cbcd1e890d`；原始任务、配置、评测器和上游 Golden 保存在本目录。
- 任务路径：`tasks/finalpool/canvas-faculty-workload-review`。
- 数据库来源：`db/init.sql.gz`，源文件 SHA256 `bd9c45b5ec9072b9f301b7df5b3e9ac65540ffcd53485329ebea44b0c3710da4`。仓库为 Apache-2.0，完整许可证保存在 `LICENSE-Cowork-Bench`。
- 上游资产是合成教育数据；投影只保留教师授课关系、课程、学生选课计数所需的学生 ID、作业 ID 和必要状态，排除学生姓名、邮箱、成绩等无关字段。

## 输入和证据链

| 证据 | 本地路径 | 作用 | 变换 |
|---|---|---|---|
| 上游教师目录 | `data/input_files/faculty_directory.csv` | 机构背景和部门语境 | 原字节复制，SHA256 `c799754e494d9edbc33992852f1aec9c771b1e15dd545b20d1cd64e4f1f19476` |
| 上游政策 | `data/input_files/workload_policy.pdf` | 4 门课程、40 小时、0.5 小时/学生、16 周等规则 | 原字节复制，SHA256 `8def78ba5d60a69ce6fa88590562dc030bf98cdf0b1d470d85a478baad52ccbc` |
| 固定规范 API 快照 | `data/input_files/workload_standards.json` | 机器可读阈值 | 原字节复制，SHA256 `a02b95499e815bcfad690616613f9ccf45db3163af7292e3cd74894cbddf9127` |
| `canvas.courses` 投影 | `data/input_files/canvas_courses.csv` | 课程代码、主题、学生数和状态 | `extract_canvas_snapshot.py` 从源 SQL 提取；主题按可见代码前缀 AAA–GGG 映射为 Analytics、Biochemistry、Computing、Design、Economics、Finance、Governance |
| `TeacherEnrollment` 投影 | `data/input_files/canvas_teacher_enrollments.csv` | 教师与课程的一对多授课关系 | 仅保留教师 ID、Canvas 显示名、Login_ID、课程 ID/代码 |
| `StudentEnrollment` 投影 | `data/input_files/canvas_student_enrollments.csv` | 按课程逐行重算学生数 | 仅保留课程 ID、学生 ID、选课状态；不保留学生姓名/邮箱/成绩 |
| 作业投影 | `data/input_files/canvas_assignments.csv` | 按课程逐行重算作业数 | 仅保留作业 ID、课程 ID、名称 |

`canvas_snapshot_manifest.json` 记录了 33 名教师、41 条教师授课关系、23 门课程、32,605 条学生选课记录和 206 条作业记录。输入投影与 `data/`、`environment/input_files/` 字节一致，所有哈希见 `metadata/input-manifest.json`。

## 独立真值和改造边界

参考解法和 Judge 都从四个 CSV 重新聚合，不读取上游 Golden。教师行按 `Login_ID` 字母顺序排列，保持 Canvas 显示姓名；课程数按教师的不同 `Course_ID` 计数；学生数和作业数按每条教师-课程关系累加。部门汇总的 `Course_Count` 是不同课程代码数，`Total_Students` 是教师-课程授课关系学生负荷之和，因此同一课程有多名教师时会分别计入，这是对上游 SQL 语义的显式化。

`Est_Grading_Hours` 和 `Weekly_Grading_Hours` 保留一位小数，采用 Excel `ROUND` 的半舍入规则（例如 1364/16 = 85.3）；上游 Golden 对 85.25 和 145.25 使用了不同的半偶舍入，Judge 以政策计算和适配规则为准并保留该差异供人工复核。由输入独立重算的结果为 33 名教师、23 名超负荷教师，部门汇总为 7 行。

原任务要求写 Google Sheet 并通过 Canvas/HTTP API 获取状态。本适配把同一固定状态投影为本地 CSV，把共享表替换为可编辑的 `Faculty_Workload_Analysis.xlsx`，并保留 `Faculty_Workload.xlsx` 和 `Workload_Review.pptx` 三个交付物。没有新增外部邮箱、日期或服务状态，也没有把截图作为表格或演示文稿内容。

## 要求到检查器映射

| 要求 | 证据 | Judge |
|---|---|---|
| 教师负荷逐人统计 | Canvas 四个投影 CSV + workload 标准 | R002 |
| 学科汇总 | 课程代码主题映射 + TeacherEnrollment | R003 |
| 本地共享分析表 | 原任务 Google Sheet 意图，改为 `Faculty_Workload_Analysis.xlsx` | R004 |
| 演示文稿总数、部门表和建议 | 原任务 PPT 意图 + 独立聚合结果 | R005/R006 |
| 输入追溯和政策规则 | manifest、SHA256、政策 PDF、规范 JSON | R007 |
| 原生可编辑交付 | OOXML、Excel 表格/冻结窗格、PPT 文本和表格 | R001/R008 |

## 验证记录

- `solution/solve.sh` 生成 Golden；Golden Judge：`1.000000`。
- fresh replay（独立临时输出目录）：`1.000000`。
- 空目录、缺少 PPTX、损坏 PPTX：`0.000000`，并干净退出。
- 修改一个教师课程数：`0.660000`；把 PPT 超负荷总数改为 99：`0.640000`；额外交付文件：`0.880000` 且因 R001 hurdle 不通过。
- LibreOffice Calc/Impress 转 PDF 成功：`Faculty_Workload.xlsx` 4 页、`Faculty_Workload_Analysis.xlsx` 3 页、`Workload_Review.pptx` 4 页；已检查标题页、概览、部门表、建议页以及首个 Excel 页面，无空白/损坏或文本重叠。
- 全库静态检查和 `git diff --check` 在本题收口后重新执行。

## 难度、泄漏与待审

- 作者假设：L3/advanced。任务需要读取 PDF/JSON/多份 CSV，理解教师-课程多对多负荷，完成两个一致的 Excel 表和一份 PPT；预估熟练办公人员约 60 分钟。
- leakage group：`cowork-faculty-workload`；与 Cowork-Bench 中依赖同一 Canvas 初始化快照的教师/课程分析题隔离。
- 待人工决定：Canvas SQL 快照中学生 ID 虽为合成数据但是否适合最终发布；按授课关系累加部门学生负荷的语义；本地共享表替代 Google Sheet 的保真度；上游半偶舍入差异；最终难度与是否纳入正式题库。
