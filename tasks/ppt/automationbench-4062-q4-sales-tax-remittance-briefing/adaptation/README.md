# AutomationBench 4062 逐题适配记录

本题直接阅读 `finance.sales_tax_remittance` 的 Query、完整 initial state 和 assertions，并逐字段静态化 8 行 Q4 revenue、3 行 tax rates、政策邮件和 NY exemption 邮件。没有新编州、月份、税率或销售额。

制题审计发现上游邮件说 exemption 金额是 `approximately $30,000` / `$45,000`，但断言据此精确计算 NY `$16,800`。本适配不隐瞒：Golden 保留 `≈`，同时说明采用上游 asserted arithmetic。Judge 独立从输入重算 CA、NY、OR 月度与州级金额，并检查“只准备、不得申报、不得付款”的授权边界。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。来源漏行/错数、CA 生效税率错、NY exemption 错、月度税额错、州图/汇总错、遗漏 approximately 边界、违规申报/付款、错收件人、缺来源、压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `9147f69684bac9c51d73d459cabdfd4c0ea536f15b2db29441856bf4a44a2cfc`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4062-oracle-sales-tax-20260829`，trial 为 `automationbench-4062-q4-sales-ta__jqMhBEs`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已逐页目视；来源表、月度 schedule、政策例外、州图和授权页均无重叠裁切，`visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 模型逐题读取并检查了任务指令、配置、确定性 Rubric、任务专属 Judge、负例预期、原适配记录、上下文 JSON、两封政策/证书 EML，以及 Excel `Q4 2025 Revenue` 八行与 `Tax Rates` 三行真实数据。
- Golden `Q4_2025_Sales_Tax_Remittance_Briefing.pptx` 共 6 页，已先查看 3×2 接触表，再逐页查看 2000×1125 原图：审阅结论、来源登记、月度 schedule、政策与例外、州级 liability 图/表、受控交付流程。
- 该题是从冻结税务材料开放式生成审阅简报，可采用多种专业表格和图表布局，故使用 `reference_free`；Golden 仅用于制定任务特定锚点和校准。
- `V001` 检查应缴总额与不申报/不付款边界，`V002` 检查来源表到 schedule 的数据证据链，`V003` 检查三类州级证据与 approximately caveat，`V004` 检查 liability 图表/汇总表的视觉解释，`V005` 检查四步受控交付流程，`V006` 检查税务审阅语境下的克制配色和六页一致性；不重复确定性 Judge 已覆盖的精确数字、事实和文件结构。
- Golden 逐页检查未发现明显遮挡、裁切或不可读区域；第 5 页图表数值采用无千分位格式，属于轻微专业呈现风险，VLM 应结合实际可读性如实判断。
- Golden 已用固定 Codex CLI 0.148.0、the configured VLM provider 连续真实评分 5 次，visual reward 为 `1.000000 / 0.981481 / 1.000000 / 1.000000 / 0.981481`，均值 `0.992593`、总体标准差 `0.009072`（小于 0.05）；确定性分均为 `1.0`，70/30 综合分均值为 `0.997778`。
- 任务特定中等负例只把第 4 页纽约 `CERTIFICATE` 证据卡整体左移 2.35 英寸，使其明显叠压加州 `RATE CHANGE` 卡，破坏 CA/NY/OR 三类证据的州别并列与完整读取；6 页仍可渲染，全部数字和文本仍在文件中，确定性 Judge 仍为 `1.0`。真实 VLM visual reward 降至 `0.848765`，70/30 综合分为 `0.954630`，严格低于 Golden 均值且高于空白或损坏交付。
- 评分证据保存在 `.coworkevo/vlm-calibration/ppt/automationbench-4062-q4-sales-tax-remittance-briefing/`，不提交评测 artifacts。
