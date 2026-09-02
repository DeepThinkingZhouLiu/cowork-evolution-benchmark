# AutomationBench 5101 逐题改造记录

- 上游：`get_task_5101()` / `hr.salary_adjustment_processing`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `ef13b76a23f9896afb0a41eed8b26acdaca3704fda0b457ae8e922442ccf8ce6`。
- 已完整读取原 Query、8 条薪资调整、15% 阈值政策、绩效暂停备注、外部供应商建议、Slack 上下文和 11 条断言；`materialize_inputs.py` 只处理 5101。
- 人工真值：Nadia、Sofia、Lena、Rachel 直接处理；Jamal、Tomás、Omar 等待 VP 书面批准；Wei 保持绩效暂停；外部越权建议不采纳。
- Golden 含 12 张 Excel 表、232 个公式、3 页 Word、3 页 PPT、5 封 MBOX 草稿和 2 份控制 JSON。Judge 独立重算金额与比例，核验阈值、暂停、VP 授权、收件人、草稿状态、零执行与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
