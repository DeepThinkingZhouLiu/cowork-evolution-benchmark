# AutomationBench 5114 逐题改造记录

- 上游：`get_task_5114()` / `hr.termination_checklist`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `1577b21f2f048bc6a1e69e47c2aa8eba81dea0bed08ca992269ff9aec6bdeece`。
- 已完整读取原 Query、9 行台账、HR 终止程序、Slack 上下文和 10 条断言；`materialize_inputs.py` 只处理 5114。
- 人工真值：7 人准备清单草稿；Dina 已处理、Isaac 已撤回；Bianca、Ethan、Holly 的 PTO 使用未知措辞；HR 只请求 Payroll 计算，不发放薪资。
- Golden 含 12 张 Excel 表、207 个公式、3 页 Word、3 页 PPT、21 封 MBOX 草稿和 3 份控制 JSON。Judge 独立核验人员、PTO、收件人、跳过、五类草稿、零副作用与 Office 安全。
- 12 类负例全部 0 分；两次 fresh replay 与两次 Harbor Docker Oracle 均 1.0、0 异常；18 页目视检查通过。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格口径完整保存在 `rubric.v1.json`、`tests/judge_v1.py` 和 `tests/negative_cases_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 按本题七名处理对象逐人核验 Process、IT、Payroll、Tracker 四类 Excel 队列，按21封邮件逐封核验 To、主题、类型、日期、PTO 与未执行声明，并分别核验 Dina 已处理、Isaac 撤回、三条未知 PTO、七条 Slack/Tracker 草稿和17项 Execution 控制。仅核心交付缺失、为空、损坏、不可解析或只交原样输入时硬归零。
- v2 Golden 与 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。28 个 v2 负例中 16 个核心失败为 `0`，其余12个任务语义错误为 `0.725294–0.997143`；局部字段错误保留高部分分，误发薪、误撤权、误处理撤回对象等高风险错误明显更低。
- `Termination_Checklist_Review.pptx` 使用独立五项视觉 Rubric，分别检查9/7/1/1范围、七人日期与 PTO、Bianca/Ethan/Holly 未知余额、21/7/7草稿与零执行、跨页可读性和治理叙事。Codex CLI `0.148.0` + `your-vision-model` 的五次 Golden 视觉分为 `0.972/0.938/0.9469/0.9635/1.0`，均值 `0.96408`、总体标准差 `0.021581`。
- 任务专属视觉负例仍可打开且确定性分为 `1.0`，但第2页整列 Payroll PTO 被遮挡后，视觉分降至 `0.517`，70/30 组合分为 `0.8551`。六次视觉调用均成功，未解决基础设施失败为 `0`。
- 仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，未伪造 Harness 结果。
