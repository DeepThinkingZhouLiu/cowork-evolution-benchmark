# AutomationBench 5057 逐题改造记录

- 上游：`get_hr_multi_app_new_hire_setup_task()` / `hr.multi_app_new_hire_setup`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `d7431ffb15bfc6b6c4c60b74992011d3aee221495faf4e35bf620ecdd5f0d88a`。
- 已完整读取原 Query、两位新人的七个业务字段、应用初始状态和 11 条断言；`materialize_inputs.py` 只处理 5057。
- 人工真值：Alicia 与 Tyrone 各需 BambooHR/Drive/Jira/Gmail 一条草稿；Drive 父目录为 `folder_employees`；David Okonkwo 和 Lisa Wang 仅是经理，新人档案和欢迎邮件均为0。
- Golden 含 12 张 Excel 表、32 个公式、3 页 Word、3 页 PPT、6 条系统动作 JSON、2 封 MBOX 草稿和 1 份控制 JSON。Judge 独立核验身份、字段、父目录、四系统对账、经理排除、零执行与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；18 页目视检查通过。
