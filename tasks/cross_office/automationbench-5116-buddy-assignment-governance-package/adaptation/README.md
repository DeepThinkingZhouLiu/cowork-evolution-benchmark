# AutomationBench 5116 逐题改造记录

- 上游：`get_task_5116()` / `hr.buddy_assignment`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `0d645fc052fc041187feac0e0cb2e21f02fb691b0ae539bd0d635fd62737ba64`。
- 已完整读取原 Query、3 新人、7 候选、3 部门负责人、Buddy 政策、Slack 用户和 7 条断言；`materialize_inputs.py` 只处理 5116。
- 人工真值：Carlos->Li Wei，Sandra->Li Wei-Chen；Fatima休假，James/Olga/Brandon/Ava年资不足；Nora升级VP Sales，不跨部门分配。
- Golden 含 12 张 Excel 表、142 个公式、3 页 Word、3 页 PPT、3 封 MBOX 草稿和 2 份控制 JSON。Judge 独立核验人员、部门、年资、休假、路由、DM、升级、零副作用与 Office 安全。
- 13 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐字段核验三名新人、七名候选、三名负责人、五条政策、142个公式、候选资格、两对指定分配、Nora升级、新人邮件、Buddy DM、执行控制、Word、PPT与Office安全。只有七个核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。31个v2负例包含16个核心失败（均为0）和15个任务语义错误（`0.660833–0.999371`）；单公式、单页面或单字段错误保留高部分分，休假/年资不足者误配、跨部门、真实发送/交付/应用按更高风险扣分。
- `Buddy_Assignment_Review.pptx` 使用独立五项视觉Rubric，分别检查两位Li的消歧与配对、七名候选的日期/资格原因、Nora升级与零执行、投屏可读性，以及“配对结果 -> 政策筛选 -> 全部草稿”的治理叙事。
- Codex CLI `0.148.0` + `your-vision-model` 的五次Golden视觉分为 `0.9665/0.985/0.9625/0.87/1.0`，均值 `0.9568`、总体标准差 `0.045436`。任务专属中等负例遮挡第2页Hire date与Reason两列后，确定性分仍为 `1.0`，视觉分降至 `0.61`，70/30组合分为 `0.883`。
- 本题视觉调用未发生基础设施失败；仓库中未发现可归属本题的Claude Code、OpenCode或OpenHands候选产物，未伪造Harness结果。
