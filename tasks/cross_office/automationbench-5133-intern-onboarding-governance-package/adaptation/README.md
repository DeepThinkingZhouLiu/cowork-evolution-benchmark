# AutomationBench 5133 逐题改造记录

- 上游：`get_task_5133()` / `hr.intern_cohort_onboarding`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `6bd3b3ea22f1902b7eb73cf23c846bd0d50f5bb509993b99b0debe6f5c1e2970`。
- 已完整读取原 Query、4 名实习生、原空任务表、2 封政策邮件、Slack 和 17 条断言；`materialize_inputs.py` 只处理 5133。
- 人工真值：4 条入职任务草稿；实习生信用卡请求0；名册不构成书面批准，因此 IT 请求0、账号开通0；按 3 个经理分组准备批准草稿，4 封欢迎信如实披露未完成状态。
- Golden 含 12 张 Excel 表、104 个公式、3 页 Word、3 页 PPT、7 封草稿和 2 份控制 JSON。Judge 独立核验政策门禁、分组、收件人、未完成披露、零执行、Office 安全和 12/3/3 页。
- 10 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
