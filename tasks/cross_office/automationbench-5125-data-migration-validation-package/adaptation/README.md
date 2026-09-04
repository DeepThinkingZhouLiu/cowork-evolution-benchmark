# AutomationBench 5125 逐题改造记录

- 上游：`get_task_5125()` / `hr.data_migration_validation`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `67dd60750067659f79f1fec639c84bcf7a242eb871f0d7052bb08458521c5b9b`。
- 已完整读取原 Query、8 条迁移记录、更正邮件、同名陷阱和 7 条断言；`materialize_inputs.py` 只处理 5125。
- 人工真值：EMP-6601 职级修正为 Senior Engineer；EMP-6618 部门修正为 Product；EMP-6655 是另一位 Michael Park，Sales / Sales Rep 记录正确；其余五条也为 Validated。
- Golden 含 10 张 Excel 表、104 个公式、3 页 Word、3 页 PPT、1 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验 ID 消歧、差异、状态草稿、邮件边界、零更新与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 均 1.0 且交付逐字节一致；16 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `1cb718b3b70826012753a52b36b829a4db3ebfefcce775e55b31aefc0bc2e47e`、`7075873342f7fcc018e36d92a61a33f5d0b85bab16b046339f7629525f9c8c78`、`6af1692906eae61e6115c53b0bea7ef37a5375b2b1c1ab84840ee90be36f5d22`；默认v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2逐项核验8条源记录、更正邮件与4条证据、全部104个任务专属公式、2条差异、6条Validated、HRIS邮件、8条状态草稿、两份控制JSON、Word/PPT和跨文件一致性。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden与fresh replay均为`1.0`，v1 Golden仍为`1.0`；fresh七份产物与Golden逐字节一致。32个v2负例中16个核心失败均为0，16个任务语义错误为`0.628081–0.998947`；单字段错误连续扣分，误报EMP-6655、遗漏权威更正、虚假应用/发送/源更新按风险显著降分。
- `Data_Migration_Validation_Review.pptx`的独立五项视觉Rubric检查第1页8/2/6/0与两条更正、第2页两个Michael Park依Employee ID的并列消歧、第3页8/1草稿与全部0执行，并检查长ID/状态可读性和“结果 -> 消歧 -> 离线边界”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra`五次Golden视觉分为`[0.973, 1.000, 0.9754, 1.000, 1.000]`，均值`0.98968`、总体标准差`0.012662`。任务专属中等负例保留第2页两个Employee ID和原生表格，只遮挡Department、Title和Decision正文，使两位Michael的身份差异无法肉眼复核；确定性分`1.0`、视觉分`0.595`、70/30组合分`0.8785`。
- 五次Golden和一次视觉负例调用均正常完成，基础设施失败为0；失败时必须标记infra failure而不是0分或静默退化。仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，未捏造Harness结果。
