# DOCX 60 题 Judge 审计记录（2026-09-07）

本轮逐题阅读了 instruction、真实输入、Golden、rubric 和 `judge.py`。审计只处理确定性 DOCX 评分，不引入 VLM，也不将视觉审美纳入本轮得分。

## 审计结论

- 60 道题全部完成逐题审计。
- 24 道现有 Judge 已经使用输入真值、结构或渲染检查，本轮保留。
- 36 道修订 Judge/Rubric：核心事实改为必过项，取消“文件非空即可过”的漏洞；其中错误 Golden 一并修复。
- `officebench-3-84-sleep-analysis` 的上游指令要求睡眠时长和体重，但发布工作簿没有这两列。v0.2 保留原始来源记录，将可执行任务修复为“满意度与年龄、收入、性别的关系”，禁止伪造缺失字段。
- 评分变更使用新的 rubric/Judge 版本和 task `0.2.0`；旧结果不能与新结果混用。

## 逐题结果

| #  | 任务                                                        | 处理       | 核心判断 |
|---:|-------------------------------------------------------------|------------|----------|
| 01 | `cowork-bench-rzd-trip-plan-from-outline`                    | 保留       | 车次、时刻、日程、预算和文档结构已有明确真值与门槛 |
| 02 | `forte-legal-020-power-of-attorney-risk-report`              | 保留       | 六份委托书、风险集合、等级、引用和表格已逐项核对 |
| 03 | `llm-docx-d3-plain-english-closing-conditions`               | 保留       | 五项交割条件、通俗化和非目标章节保留均有硬门槛 |
| 04 | `llm-docx-f6-accept-all-redlines`                            | 保留       | 修订接受、文本真值、OOXML 修订痕迹和渲染已检查 |
| 05 | `llm-docx-s4-revise-memo-meeting-notes`                     | 保留       | 备忘录事实、会议纪要修订和非目标内容保留已检查 |
| 06 | `officebench-1-12-class-roster-counts`                       | 保留       | 统计值由输入工作簿现场重算 |
| 07 | `officebench-1-12-final-pass-count`                          | 保留       | 通过人数由成绩工作簿现场重算 |
| 08 | `officebench-1-12-xihozun-hw2`                               | 保留       | 指定学生与作业成绩由输入行精确推导 |
| 09 | `officebench-1-13-cs-transfer-eligibility`                   | 保留       | 转专业资格按输入课程和规则重算 |
| 10 | `officebench-1-13-cs161-prerequisites`                       | 保留       | 先修课资格按输入学生记录重算 |
| 11 | `officebench-1-15-house-creak`                               | 修订 v0.2 | 原文段落必须完全保留，指定句必须是新的末段 |
| 12 | `officebench-1-15-project-report-reference`                  | 修订 v0.2 | 项目报告原文必须不变，完整书目条目必须作为末段 |
| 13 | `officebench-1-15-syllabus-integrity`                        | 修订 v0.2 | 原教学大纲必须是前缀，学术诚信文本必须完整且有序追加 |
| 14 | `officebench-1-16-random-paragraph`                          | 修订 v0.2 | TXT 全文必须不改动地成为唯一内容段 |
| 15 | `officebench-1-16-test-sentence`                             | 修订 v0.2 | 唯一非空段必须与目标句逐字相等 |
| 16 | `officebench-1-19-agreement-terms`                           | 修订 v0.2 | 双方及角色、四项条款和生效日期均必须与 PDF 一致 |
| 17 | `officebench-1-19-paper-abstract`                            | 修订 v0.2 | 摘要核心句完整，且不能夹带作者、图注和脚注噪声 |
| 18 | `officebench-1-19-paper-metadata`                            | 修订 v0.2 | 论文标题和八名作者均为必要真值，禁止整篇 PDF 粘贴 |
| 19 | `officebench-1-3-rename-docx`                                | 修订 v0.2 | 重命名后 DOCX 必须与源文件逐字节一致 |
| 20 | `officebench-2-18-highest-score-statement`                   | 修订 v0.2 | 从成绩表重算唯一最高分 Rachel=100，不得写入其他学生 |
| 21 | `officebench-2-18-homework-assignment`                       | 修订 v0.2 | 指定作业句只能出现一次，仅保留标题和正文 |
| 22 | `officebench-2-18-shopping-list`                             | 修订 v0.2 | 表格的表头、数值、行数和顺序逐行与 XLSX 匹配 |
| 23 | `officebench-2-19-teaching-report`                           | 保留       | 学生数、成绩统计和分布由输入重算，已有门槛 |
| 24 | `officebench-2-2-concert-address`                            | 修订 v0.2 | 只提取 PDF 中的表演者和地址，标签与可编辑段落齐全 |
| 25 | `officebench-2-2-concert-contact`                            | 修订 v0.2 | 表演者、日期、时间、场地、地址和票务 URL 全部精确 |
| 26 | `officebench-2-20-all-cs161`                                | 修订 v0.2 | 名单从工作簿重算，逐行、顺序、格式和行数全部精确 |
| 27 | `officebench-2-20-cs-major-list`                             | 修订 v0.2 | 仅输出 Major=CS 的学生，逐行与输入顺序精确匹配 |
| 28 | `officebench-2-20-final-perfect`                             | 修订 v0.2 | 仅输出 Final=100 的学生，不接受多写、少写或乱序 |
| 29 | `officebench-2-21-bill-courses`                              | 修订 v0.2 | 课程列数值和原始顺序精确，不得混入其他列 |
| 30 | `officebench-2-22-house-index-report`                        | 修订 v0.2 | 要求关键年份/指数、历史趋势、2030 预测和明确假设 |
| 31 | `officebench-2-22-midterm-report`                            | 保留       | 人数、两次平均分、差值和原生表格已由输入重算 |
| 32 | `officebench-2-26-invitation-text`                           | 修订 v0.2 | PDF 与 DOCX 文本归一化后必须全文一致 |
| 33 | `officebench-2-26-market-analysis-text`                      | 修订 v0.2 | PDF 与 DOCX 文本归一化后必须全文一致 |
| 34 | `officebench-2-30-rename-student-id`                         | 修订 v0.2 | 文件名为识别出的学号，段落与顺序不得改变 |
| 35 | `officebench-2-36-grade-report`                              | 修订 v0.2 | OCR 必须包含学生、GPA、学分和大量具体课程/成绩锚点 |
| 36 | `officebench-2-37-event-frequency`                           | 保留       | 日历事件频数、总数、排名和唯一最大值由 ICS 重算 |
| 37 | `officebench-2-39-all-emails`                               | 保留       | 四封邮件逐字段匹配，且按 `Sent at` 时间排序 |
| 38 | `officebench-2-39-earliest-email`                            | 修订 v0.2 | 修正为按正文 `Sent at` 比较，必须命中 2022 年租金提醒邮件 |
| 39 | `officebench-2-39-latest-email`                              | 修订 v0.2 | 修正为按正文 `Sent at` 比较，必须命中 2024 年申请批准邮件 |
| 40 | `officebench-2-6-homework-ocr`                               | 修订 v0.2 | OCR 必须包含课程、作业号、学生、学号、截止日期和题目主题 |
| 41 | `officebench-2-7-paper-summary`                              | 保留       | Transformer 的结构、复杂度、训练配置、结果和局限已分组核对 |
| 42 | `officebench-3-22-homework-folder`                           | 修订 v0.3 | 两份交付必须一致，且 OCR 内容必须命中具体作业锚点 |
| 43 | `officebench-3-56-class-homework`                            | 修订 v0.2 | 必须同时写明 Class 1 和指定作业，不允许泛化作业文档 |
| 44 | `officebench-3-58-shopping-docx`                             | 修订 v0.2 | 原生 Word 表格的表头、全部数据行和顺序与 XLSX 精确一致 |
| 45 | `officebench-3-60-cs161-folder`                              | 修订 v0.2 | 指定子目录中的 123 行名单必须与工作簿精确一致 |
| 46 | `officebench-3-63-score-report-folder`                       | 修订 v0.2 | 指定子目录中报告必须给出 20、80.05 和 78.30 三项真值 |
| 47 | `officebench-3-77-house-price-prediction`                    | 修订 v0.2 | 图像输入必须导出合理的 2030 数值、历史趋势和风险假设 |
| 48 | `officebench-3-8-progress-count`                             | 修订 v0.2 | 必须明确给出“进步学生为 5 人”，不再接受任意数字 |
| 49 | `officebench-3-82-top-five`                                 | 修订 v0.2 | 前五名姓名、分数、并列次序和行数都必须精确 |
| 50 | `officebench-3-84-sleep-analysis`                            | 修订 v0.2 | 修复上游字段错配，重算 30 行满意度的均值、相关和性别分组 |
| 51 | `officebench-3-91-calendar-conflicts`                        | 修订 v0.2 | 必须列出五组重叠及精确交叉时段，并排除 14:00 端点相接 |
| 52 | `officeval-003-meeting-content-table-ruling`                 | 保留       | 会议记录行列真值、边框和原生表格结构已精确检查 |
| 53 | `officeval-006-pdf-to-editable-math-template`                | 保留       | 21 道题、答案、解析、OMML、图示、页脚和五页渲染已检查 |
| 54 | `officeval-010-editable-technical-roadmap-reconstruction`    | 保留       | 流程图节点、连线、可编辑性和版面结构已检查 |
| 55 | `officeval-018-report-formatting-fidelity`                   | 保留       | 报告文本、样式、页面几何和渲染保真已检查 |
| 56 | `officeval-022-format-transfer-cover-directory`              | 保留       | 封面、目录、样式迁移、字段和渲染已检查 |
| 57 | `officeval-029-image-to-editable-training-form`              | 保留       | 表单字段、顺序、表格、可编辑性和渲染已检查 |
| 58 | `officeval-033-retail-price-catalog-uplift`                  | 保留       | 零售价 1.1 倍计算、全表数值和格式保留已检查 |
| 59 | `officeval-036-pdf-statement-to-word-table`                  | 保留       | PDF 流水、金额、顺序、原生表格和页面渲染已检查 |
| 60 | `officeval-040-advanced-formatting-and-watermark`            | 保留       | 水印、页面布局、段落样式、内容保留和渲染已检查 |

## 验证要求

- 修订题的 Golden 必须通过新 Judge。
- 每道题都要拒绝“可打开但核心答案错误”的 DOCX 负例。
- OCR 题必须在题目 Docker 环境中执行 Oracle，不以宿主机缺少 Tesseract 为替代。
- 所有 rubric JSON、Judge Python、项目单测和公开任务协议校验必须通过。
