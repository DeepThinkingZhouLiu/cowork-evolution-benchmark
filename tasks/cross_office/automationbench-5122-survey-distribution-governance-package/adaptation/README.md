# AutomationBench 5122 逐题改造记录

- 上游：`get_task_5122()` / `hr.survey_distribution`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `eee29f4138f088b3a7154009b6704de12af8c934e44ed011dc842d7134ef7bff`。
- 已完整读取原 Query、10 人名单、3 封政策、Slack 频道和 8 条断言；`materialize_inputs.py` 只处理 5122。
- 人工真值：6 人合格；Anna、Clara、Derek、Iris、Jake 可准备邀请；Elena 合格但缺邮箱；Boris/Faisal 因精确 Legal 排除，Clara 的 Legal Operations 保留，Grace/Hugo 分别因合同工/休假排除。
- Golden 含 10 张 Excel 表、140 个公式、3 页 Word、3 页 PPT、5 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验授权优先级、精确部门、缺口、路由、零副作用与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；16 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。三份 v1 归档分别与改造前 HEAD 逐字节一致。
- v2 逐字段核验10人名单、3封政策、证据哈希、140个公式、6/5/4/1资格与投递分流、5封邮件草稿、All-hands草稿、Elena缺口、Execution、Word/PPT和七交付一致性。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`；Fresh 的七份产物与 Golden 逐字节一致。32个v2负例包含16个核心失败（均为0）和16个任务语义错误（`0.627273–0.999220`）；单一来源、公式、收件人、缺口owner、Word/PPT事实错误保留连续部分分，邀请被排除人员、模糊Legal匹配和虚假发送/发布按风险显著降分。
- `Survey_Distribution_Review.pptx` 使用独立五项视觉Rubric，分别检查第1页10/6/5/1汇总与十人分流、第2页精确Legal边界和Elena缺口、第3页投递草稿与零执行、密集原生表格可读性，以及“总体分流 -> 政策边界 -> 离线控制”叙事。
- Codex CLI `0.148.0` + `your-vision-model` 的五次Golden视觉分为 `1.000/1.000/1.000/1.000/0.972`，均值 `0.9944`、总体标准差 `0.0112`。任务专属中等负例不改底层原生表格，仅用白色遮罩覆盖第2页Decision/Reason正文；确定性分仍为 `1.0`，视觉分降至 `0.51`，70/30组合分为 `0.853`。
- 本题五次Golden和一次负例调用均最终正常完成，基础设施失败为0；首次Golden调用因供应端并发慢响应自动重试后成功，不计失败。仓库中未找到可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
