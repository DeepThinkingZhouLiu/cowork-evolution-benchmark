# AutomationBench 1115 逐题适配记录

本题直接阅读了 AutomationBench `marketing.competitive_response_plan` 的原 Prompt、完整 initial state 和 assertions，没有用通用生成器批量编题。上游 Google Sheets 的 6 行 feature、分析师邮件、价格政策邮件和 3 条 Slack 消息被分别静态化，以保留源边界。

Golden 是人工逐页设计的 6 页 16:9 竞争响应简报：封面结论、原生 gap-status 图表、6 行原生 feature 矩阵、分析师/Slack 分源证据、响应优先级和价格治理/交付页。不执行外部邮件、HubSpot 或价格页写入，但把原断言中的 brief 事实、ticket 主题和禁止直改要求转为可确定检查的演示交付。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面配色变体均为 1.0。图表分布错误、矩阵漏行/gap 错、分析师或 Slack 事实错误、响应优先级错误、价格治理违规、受众/ticket 错、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `f4ef3ab496a253a8d57e458beb3b42ec80f1ff69cfa998b0498fd21aee9ab49b`。单题结构、公开资产、任务元数据、来源和 Rubric Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench1115-oracle-competitive-20260829`，trial 为 `automationbench-1115-stratacorp__JRwvF3P`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已渲染为接触表并逐页目视；原生 gap 图、6 行 feature 矩阵、双源证据、响应和治理页均无重叠裁切。`visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 模型逐题读取并检查了 `instruction.md`、`task.toml`、`rubric.json`、任务专属 `tests/judge.py`、`tests/expectations.json`、上游适配记录，以及 5 个冻结输入的真实内容；其中 Excel 的 `Feature Comparison` 工作表共 6 行能力记录，两个 EML 与 Slack JSON 均逐项核对。
- Golden `StrataCorp_v4_Competitive_Response_Brief.pptx` 共 6 页，已先查看 3×2 接触表，再逐页查看 2000×1125 原图：第 1 页竞争态势与 KPI，第 2 页 gap 图和 headline moves，第 3 页能力矩阵，第 4 页分析师/Slack 分源证据，第 5 页响应优先级与 watchlist，第 6 页价格治理和 ticket 行动。
- 本题属于从多源材料开放式生成管理层简报，存在多种合理版式与咨询叙事，因此采用 `reference_free`。Golden 仅用于制订与校准视觉标准，不送入 VLM，避免把单一参考设计误当成唯一答案。
- `V001` 针对封面结论与三组态势概览，`V002` 针对第 2–3 页从状态概览到六项能力下钻及跨页色义，`V003` 针对第 4 页分析师/Slack 分源和市场信号扫读，`V004` 针对第 5 页行动优先级与客户 watchlist，`V005` 针对第 6 页价格治理、ticket 下一步和首尾一致性；这些均补足确定性 Judge 无法判断的视觉层级、构图与专业表达，不重复其事实和文件结构检查。
- Golden 已用固定 Codex CLI 0.148.0、the configured VLM provider 连续真实评分 5 次，visual reward 分别为 `0.944853 / 0.977941 / 0.944853 / 1.000000 / 1.000000`，均值 `0.973529`、总体标准差 `0.024761`（小于 0.05）；确定性分始终为 `1.0`，70/30 综合分均值为 `0.992059`。
- 任务特定中等负例只把第 4 页 `FIELD / SLACK` 面板向左错移 2.65 英寸，使其与 `ANALYST NOTE` 的 `4.2/5`、`12 vs 30+` 和价格证据发生肉眼可见的叠压；文件仍可打开、6 页均可渲染且确定性 Judge 仍为 `1.0`。真实 VLM visual reward 降至 `0.766176`，70/30 综合分为 `0.929853`，严格低于 Golden 均值且高于空白或损坏交付。
- 评分证据保存在 `.coworkevo/vlm-calibration/ppt/automationbench-1115-stratacorp-competitive-response-deck/`；该目录只保存本地评测产物，不提交仓库。Golden 的稳定扣分点是第 2 页柱体未复用状态色，以及第 3 页 Priority 列略窄；它们作为参考答案视觉风险保留记录，不通过修改 Golden 或放宽 Rubric 掩盖。
