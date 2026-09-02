# AutomationBench 5122 逐题改造记录

- 上游：`get_task_5122()` / `hr.survey_distribution`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `eee29f4138f088b3a7154009b6704de12af8c934e44ed011dc842d7134ef7bff`。
- 已完整读取原 Query、10 人名单、3 封政策、Slack 频道和 8 条断言；`materialize_inputs.py` 只处理 5122。
- 人工真值：6 人合格；Anna、Clara、Derek、Iris、Jake 可准备邀请；Elena 合格但缺邮箱；Boris/Faisal 因精确 Legal 排除，Clara 的 Legal Operations 保留，Grace/Hugo 分别因合同工/休假排除。
- Golden 含 10 张 Excel 表、140 个公式、3 页 Word、3 页 PPT、5 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验授权优先级、精确部门、缺口、路由、零副作用与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；16 页目视检查通过。
