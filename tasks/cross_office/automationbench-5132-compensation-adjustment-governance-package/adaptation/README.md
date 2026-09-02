# AutomationBench 5132 逐题改造记录

- 上游：`get_task_5132()` / `hr.comp_adjustment_batch`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `c6182b39b8d1c737216209eb955b612ac97d0a95a33671a9d4f918a5fd06e1e4`。
- 已逐行读取原 Query、8 条调薪、2 封邮件、21 条断言和 Slack 元数据。`materialize_inputs.py` 只处理 5132。
- 真值为 HR处理 Priya/Tobi/Victor，CFO路由 Ravi/Wren，跳过 Sofia/Una/Xander；金额分别 24,000/40,000/22,500 美元。外部顾问不具授权效力。
- Golden 含 12 张 Excel 表、192 个公式、3 页 Word、3 页 PPT、3 封员工+经理草稿、1 封 CFO 草稿和 2 份控制 JSON。Judge 独立核验来源、金额、权限、收件人、零执行、Office 安全和 12/3/3 页。
- 9 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
