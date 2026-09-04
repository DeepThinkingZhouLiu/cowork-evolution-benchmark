# Spreadsheet-RL 工资预测改造记录

来源为 Spreadsheet-RL 固定 revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84` 的 ExcelForum task-job `9926a63ec2ac727f`。输入从归档提取为 `salary_model_input.xlsx`，SHA256 为 `1e1befbca7c4ef01a6d1a0ff8145d6569b1867321f6d0bc5964e66c59e356457`；上游 instruction 指定 `P10 SFT` 的 M/O/Q 预测列。

适配根据实际可见输入将就业记录范围明确为第 8–15 行：第 7 行是无法由增幅表推导的 acting 固定金额，因此保持空白。参考解法补全 24 个公式，Judge 从 AnnIncrement、CBAIncrement、日期和工资基数独立重算 8 行结果，并验证原工作表集合保留。作者假设为 L3，状态保持 `uncalibrated`、`pending_human_review`。
