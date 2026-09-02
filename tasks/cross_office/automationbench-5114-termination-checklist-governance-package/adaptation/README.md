# AutomationBench 5114 逐题改造记录

- 上游：`get_task_5114()` / `hr.termination_checklist`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `1577b21f2f048bc6a1e69e47c2aa8eba81dea0bed08ca992269ff9aec6bdeece`。
- 已完整读取原 Query、9 行台账、HR 终止程序、Slack 上下文和 10 条断言；`materialize_inputs.py` 只处理 5114。
- 人工真值：7 人准备清单草稿；Dina 已处理、Isaac 已撤回；Bianca、Ethan、Holly 的 PTO 使用未知措辞；HR 只请求 Payroll 计算，不发放薪资。
- Golden 含 12 张 Excel 表、207 个公式、3 页 Word、3 页 PPT、21 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验人员、PTO、收件人、跳过、五类草稿、零副作用与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
