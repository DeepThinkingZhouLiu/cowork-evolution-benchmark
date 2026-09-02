# AutomationBench 5124 逐题改造记录

- 上游：`get_task_5124()` / `hr.mandatory_meeting_scheduling`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `3eb32244f6d1290f32b8786f67e19237253e94bd03ecfcc9bd220c582a44dd1a`。
- 已完整读取原 Query、8 人名单、豁免/VP+/日历政策、Slack 频道和 10 条断言；`materialize_inputs.py` 只处理 5124。
- 人工真值：Alice、Carol、Dave、Femi、Grace 必到；Bob/Hugo 因 Support 在岗豁免，Elena 因批准休假豁免；Grace 的 VP+ 覆盖被显式保留；日历修改为 0。
- Golden 含 11 张 Excel 表、112 个公式、3 页 Word、3 页 PPT、8 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验出席决策、时间、收件人、录像路由、日历越权阻断、零副作用与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；17 页目视检查通过。
