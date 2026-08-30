# OfficeBench 2-19 改造记录

来源为 OfficeBench 固定提交 `b978b808667c32b52ce19a67ce1def1de9ae02b7` 的 `2-19/0`。输入 `score.xlsx` 已人工读取并改名为 `student_scores.xlsx`，SHA256 为 `e151e7aff861749d1fce7bea0e584263cdc5182df25f9b6a9a96cb62ca9b19f8`。上游只要求生成 `teaching.docx` 且 evaluator 只检查文件存在，无法作为完整真值；本适配从 20 条输入记录独立计算 12 名学生、平均分 79.35、A/B/C/D/F 分布 5/6/3/6/0，并用可复核表格和阈值化教学建议补足交付判据。

参考解法合并重复姓名，输出班级概览、成绩分布、逐学生统计和教学优先级；Judge 不读取任何 Golden，从输入重算所有数字并解析 Word 表格。作者假设为 L3，状态保持 `uncalibrated`、`pending_human_review`。由于上游没有教师课程政策，报告中的辅导阈值（平均分 <70）和表扬阈值（平均分 >=90）已在公开指令中明确，避免隐藏要求。
