# AutomationBench 5133 逐题改造记录

- 上游：`get_task_5133()` / `hr.intern_cohort_onboarding`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `6bd3b3ea22f1902b7eb73cf23c846bd0d50f5bb509993b99b0debe6f5c1e2970`。
- 已完整读取原 Query、4 名实习生、原空任务表、2 封政策邮件、Slack 和 17 条断言；`materialize_inputs.py` 只处理 5133。
- 人工真值：4 条入职任务草稿；实习生信用卡请求0；名册不构成书面批准，因此 IT 请求0、账号开通0；按 3 个经理分组准备批准草稿，4 封欢迎信如实披露未完成状态。
- Golden 含 12 张 Excel 表、104 个公式、3 页 Word、3 页 PPT、7 封草稿和 2 份控制 JSON。Judge 独立核验政策门禁、分组、收件人、未完成披露、零执行、Office 安全和 12/3/3 页。
- 10 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，SHA256 分别为 `c9ad0f249eefcc92ddd7aff829e9f77346818093549e57c82ffd0d36ef05bb3d`、`0c56fd33f924c8ed7ddbf74133b55d22640ab71f3b6c13a24ab77fd25f5434e7`、`7f777936980c18b847dbc965488d977076ea417a79ef1480642ee6d6eba017d3`；默认 v2，`SCORING_VERSION=v1|v2` 可显式切换。
- v2 逐项核验四名实习生、冻结输入与政策证据、104 个公式、四条任务草稿、四封个人邮箱欢迎草稿、Derek/Lisa/Mira 三封经理批准草稿、信用卡绝对禁令、书面批准与 IT 所有权门禁、Slack/Execution 零执行、Word/PPT 和七交付一致性。只有核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay 均为 `1.0`，v1 Golden 仍为 `1.0`；fresh 七份产物与 Golden 逐字节一致。36 个 v2 负例中 17 个核心失败均为 0，19 个局部/高风险错误为 `0.619259–0.998889`；单事实错误保留连续部分分，误发信用卡、绕过书面批准、公开泄露个人邮箱或虚假执行会显著降分，同一执行事件不重复处罚。
- `Intern_Onboarding_Review.pptx` 的独立五项视觉 Rubric 检查第 1 页四人/开始日/经理/草稿与零执行，第 2 页任务可草拟、信用卡绝对禁止、书面批准和 IT 所有权门禁，第 3 页 3/4/0/1 离线交接与真实副作用 0，并专门检查长门禁、人员标签、表格适配和“批次准备 -> 政策范围 -> 离线交接”叙事。
- Codex CLI `0.148.0` + `gpt-5.6-terra` 五个有效 Golden 视觉分为 `1.0/1.0/0.993/1.0/1.0`，均值 `0.9986`、总体标准差 `0.0028`。任务专属中等负例保留第 2 页三类动作与政策副标题，但明显遮挡 Allowed now、Owner 和 Result；确定性分 `1.0`、视觉分 `0.628`、70/30 组合分 `0.8884`。
- 五次 Golden 与中等负例的最终视觉调用均成功，没有未解决 infra failure，也没有静默退化为纯代码总分。仓库中未找到可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未捏造 Harness 结果。
