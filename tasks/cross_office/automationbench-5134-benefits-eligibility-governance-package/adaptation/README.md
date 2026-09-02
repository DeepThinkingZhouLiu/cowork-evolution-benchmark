# AutomationBench 5134 逐题改造记录

- 上游：`get_task_5134()` / `hr.benefits_enrollment_review`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `ce17cac02ab1c337f134a085460be2460c220c5b1010ce797fbc92cdf48fd321`。
- 已完整读取原 Query、7 人队列、资格政策、orientation 名册、外部请求、Slack 和 16 条断言；`materialize_inputs.py` 只处理 5134。
- 人工真值：Alex/Farida 合格；Brianna 缺年资；Carlos/Graham 缺 orientation；Darius 缺已签表单；Elena legal hold 无行动；Jordan 非名册请求拒绝。
- Golden 含 12 张 Excel 表、189 个公式、3 页 Word、3 页 PPT、6 封员工草稿和 2 份控制 JSON。Judge 独立核验条件合取、缺口、legal hold、外部拒绝、收件人、零执行和 Office 安全。
- 10 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
