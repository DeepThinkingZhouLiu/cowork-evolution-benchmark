# AutomationBench 5134 逐题改造记录

- 上游：`get_task_5134()` / `hr.benefits_enrollment_review`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `ce17cac02ab1c337f134a085460be2460c220c5b1010ce797fbc92cdf48fd321`。
- 已完整读取原 Query、7 人队列、资格政策、orientation 名册、外部请求、Slack 和 16 条断言；`materialize_inputs.py` 只处理 5134。
- 人工真值：Alex/Farida 合格；Brianna 缺年资；Carlos/Graham 缺 orientation；Darius 缺已签表单；Elena legal hold 无行动；Jordan 非名册请求拒绝。
- Golden 含 12 张 Excel 表、189 个公式、3 页 Word、3 页 PPT、6 封员工草稿和 2 份控制 JSON。Judge 独立核验条件合取、缺口、legal hold、外部拒绝、收件人、零执行和 Office 安全。
- 10 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `00eb3228924c9e130f97d106ff0682eabb9374446fed66cd4363612a1b4cf26e`、`180cd0bb05133d11296f77e627cdd23c7203941b1b9ecd5034789ef488ce9af5`、`35d4e0c82a5adca2762754edbb4cec30f94a5cd482a6c9ed1b447997adbd3ced`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2 逐项核验七名员工、三封政策/名册/外部请求、五条规则、189 个公式、2名合格/4名未满足及精确缺口、Elena legal hold、Jordan非名册拒绝、六封员工草稿、两条登记更新草稿、Slack/Execution零执行、Word/PPT和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 七份产物与 Golden 逐字节一致。36 个 v2 负例中 17 个核心失败均为 0，19 个局部/高风险错误为 `0.627429–0.998879`；单事实错误保留连续部分分，误判资格、绕过legal hold、采纳Jordan、错发草稿或虚假执行会显著降分，同一执行事件不重复处罚。
- `Benefits_Review.pptx` 的独立五项视觉 Rubric 检查第1页7/2/4/1/0总览，第2页两名合格、四名未满足及精确缺口、Elena暂停，第3页6/2/0/1离线控制与Jordan拒绝，并专门检查长缺口、人员列表、表格适配和“处置总览 -> 逐人诊断 -> 离线控制”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 五个有效 Golden 视觉分均为 `1.0`，均值 `1.0`、总体标准差 `0.0`。任务专属中等负例保留第2页七人名单与副标题，但明显遮挡 Decision 与 Missing/override；确定性分 `1.0`、视觉分 `0.5`、70/30组合分 `0.85`。
- 五次 Golden 与中等负例的最终视觉调用均成功，没有未解决 infra failure，也没有静默退化为纯代码总分。仓库中未找到可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未捏造 Harness 结果。
