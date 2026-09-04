# AutomationBench 4040 逐题适配记录

本题直接阅读了 AutomationBench `finance.financial_close_checklist` 的原 Prompt、工具列表、完整 initial state 和 7 条 assertions，没有用通用题目生成器。

上游任务原本要求 Agent 从 Google Sheets 读取 7 行 January close checklist，向 Controller/CFO 发邮件并向 `finance-alerts` 发 Slack。本 PPT 适配不执行外部写入，而是把上游应用状态静态化成一个 XLSX 和一个 JSON，要求 Agent 生成同等事实范围的管理简报。输入中的 7 个 item、姓名、邮箱、Status、Due 和唯一 Notes 均逐字段来自上游；未编造金额或额外财务事实。

Golden 是根据这两个冻结输入逐页设计的 5 页 16:9 演示文稿：封面/总体结论、原生完成度图表、7 行原生明细表、3 个未完成责任卡和面向 Controller/CFO/Slack 的沟通绪论。

## 实际验证

- 本地 Golden、ZIP 等价重封装和非语义性封面配色变体均为 1.0。统计错、图表数值错/缺失、明细漏项/负责人错、Tax 备注漏失、due 分类错、受众/频道错、编造金额、压平和答案标记均为 0；多一个调试文件为 0.94。
- Judge 还会拒绝同一文稿中并存 57.1% 和错误完成率的相互矛盾交付。fresh replay 连续 5 次均为 1.0；单题结构和输入/Golden 资产校验通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4040-oracle-close-readiness-20260829`，trial 为 `automationbench-4040-month-end-c__n7r9PeJ`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 5 页已渲染为接触表并逐页目视；原生环图、7 行明细表、3 张责任卡和沟通末页均无重叠裁切。`visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 模型逐题读取并检查了 `instruction.md`、`task.toml`、`rubric.json`、任务专属 `tests/judge.py`、`tests/expectations.json`、原适配记录、冻结上下文 JSON，以及 Excel `January 2026 Checklist` 工作表全部 7 行真实内容。
- Golden `January_2026_Close_Readiness_Briefing.pptx` 共 5 页，已先查看接触表，再逐页查看 2000×1125 原图：第 1 页总体状态，第 2 页环图/状态计数，第 3 页七行明细，第 4 页三项未完成责任卡，第 5 页沟通绪论和交付对象。
- 本题是根据冻结清单开放式制作财务管理简报，合理答案可采用不同图表和页面结构，因此选择 `reference_free`；Golden 只用于确定任务特定视觉锚点和校准，不限制候选复制其版式。
- `V001` 检查封面决策信号，`V002` 检查环图标签、状态卡和红绿风险语言，`V003` 检查七行财务表的行列追踪和风险强调，`V004` 检查三个未完成项的责任卡及紧迫度差异，`V005` 检查沟通末页的结论—责任人—交付对象闭环；确定性 Judge 已覆盖的数值、文本、结构和可编辑性不在视觉分中重复计分。
- Golden 视觉风险：第 2 页环图的 `Incomplete; 3` 与 `Complete; 4` 标签侵入图形内部且使用分号样式，降低局部可读性；该问题不会通过降低 Rubric 标准掩盖，真实 VLM 校准应如实反映。
- Golden 已用固定 Codex CLI 0.148.0、ZCloud 与 `gpt-5.6-terra` 连续真实评分 5 次，visual reward 五次均为 `0.888060`，均值 `0.888060`、总体标准差 `0.000000`；确定性分均为 `1.0`，70/30 综合分均为 `0.966418`。
- 任务特定中等负例只把第 4 页第三张 `DUE TODAY / BLOCKED` 责任卡整体左移 2.15 英寸，使其与第二张逾期卡明显叠压，破坏三项责任人的并列扫描与紧迫度层级；5 页仍可渲染、所有文本仍在包内，确定性 Judge 仍为 `1.0`。真实 VLM visual reward 降至 `0.567164`，70/30 综合分为 `0.870149`，严格低于 Golden 且高于空白或损坏交付。
- 评分证据保存在 `.coworkevo/vlm-calibration/ppt/automationbench-4040-month-end-close-readiness-deck/`，不提交渲染图片和模型响应。
