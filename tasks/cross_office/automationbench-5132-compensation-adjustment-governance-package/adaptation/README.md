# AutomationBench 5132 逐题改造记录

- 上游：`get_task_5132()` / `hr.comp_adjustment_batch`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `c6182b39b8d1c737216209eb955b612ac97d0a95a33671a9d4f918a5fd06e1e4`。
- 已逐行读取原 Query、8 条调薪、2 封邮件、21 条断言和 Slack 元数据。`materialize_inputs.py` 只处理 5132。
- 真值为 HR处理 Priya/Tobi/Victor，CFO路由 Ravi/Wren，跳过 Sofia/Una/Xander；金额分别 24,000/40,000/22,500 美元。外部顾问不具授权效力。
- Golden 含 12 张 Excel 表、192 个公式、3 页 Word、3 页 PPT、3 封员工+经理草稿、1 封 CFO 草稿和 2 份控制 JSON。Judge 独立核验来源、金额、权限、收件人、零执行、Office 安全和 12/3/3 页。
- 9 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `948d5076a13426d3c6f974ca14f439c1cb7ba5d052343ef1e29b183fdfd9b8d6`、`8a1ee096b8a91bc01331029108a2cf987b48a0fc873e8b72e87a28af740e8cee`、`0416cc705359c0856388c3917a9af54ff4c90e195e893755ff2c7b31d9b63714`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2 逐项核验八行调薪来源、CFO政策与外部顾问邮件、六条规则、四条证据、全部 192 个公式、3 人 HR 处理、2 人 CFO 路由、3 人政策跳过、三封员工+经理草稿、CFO 草稿、Slack、Execution、Word/PPT 和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 七份产物与 Golden 逐字节一致。36 个 v2 负例中 17 个核心失败均为 0，19 个局部/高风险错误为 `0.632963–0.998961`；单字段错误保留连续部分分，HR 越权处理 Ravi/Wren、忽略三类后续覆盖、错发草稿、采纳外部顾问或虚假执行会显著降分，同一执行事件不重复处罚。
- `Compensation_Review.pptx` 的独立五项视觉 Rubric 检查第 1 页 3/2/3 与 $24,000/$40,000/$22,500 总览，第 2 页八人调薪额/Decision/Control 及权限、REVERSED、待背调、HOLD 和外部干扰边界，第 3 页 3/1/3/1 草稿和零执行，并专门检查长决策、密集表格可读性和“总览 -> 政策层级 -> 执行控制”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 五个有效 Golden 视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例保留第 2 页八名员工、调薪额和副标题，但明显遮挡 Decision 与 Control；确定性分 `1.0`、视觉分 `0.5725`、70/30 组合分 `0.87175`。
- Golden 第三次首次调用因 Provider 返回 `Service temporarily unavailable` 被明确标记为一次 infra failure，未计 0 分也未退化；同配置重试成功，因此累计 1 次失败、已解决 1 次、未解决 0 次。仓库中未找到可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未捏造 Harness 结果。
