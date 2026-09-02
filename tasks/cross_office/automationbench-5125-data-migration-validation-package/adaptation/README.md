# AutomationBench 5125 逐题改造记录

- 上游：`get_task_5125()` / `hr.data_migration_validation`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `67dd60750067659f79f1fec639c84bcf7a242eb871f0d7052bb08458521c5b9b`。
- 已完整读取原 Query、8 条迁移记录、更正邮件、同名陷阱和 7 条断言；`materialize_inputs.py` 只处理 5125。
- 人工真值：EMP-6601 职级修正为 Senior Engineer；EMP-6618 部门修正为 Product；EMP-6655 是另一位 Michael Park，Sales / Sales Rep 记录正确；其余五条也为 Validated。
- Golden 含 10 张 Excel 表、104 个公式、3 页 Word、3 页 PPT、1 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验 ID 消歧、差异、状态草稿、邮件边界、零更新与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；16 页目视检查通过。
