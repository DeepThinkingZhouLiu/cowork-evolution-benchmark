# AutomationBench 4044 逐题适配记录

本题直接阅读 `finance.po_three_way_match` 的 Query、完整 initial state 和 assertions，并逐字段静态化 Invoice Batch、Purchase Orders、Receiving Reports 和 tolerance policy。没有新编供应商、金额、数量、日期或状态。

制题审计发现一个重要的上游语义缺口：政策只定义“纯价格偏差在 5% 以内”时进入 Price Variance Review，但 VI-002 的偏差是 7.14%，上游 assertion 仍指定 Review。本适配没有把这个矛盾藏起来：Golden 保留上游状态，同时要求显式标出政策缺口；Judge 两者独立检查。

## 实际验证

- Golden、ZIP 等价重封装和非语义性封面变体均为 1.0。来源行错、VI-001 边界错、VI-002 状态错、VI-003 数量/状态错、图表错、遗漏政策缺口、收件人错、矛盾批准声明、来源缺失、整页压平和答案标记均为 0；多一个调试文件为 0.98。
- 从空输出目录执行 `solve.sh -> Judge` 连续 5 次均为 1.0，输出 SHA256 稳定为 `760721adac961f2bcb16c2a4c91e1c3d13dd046bc797238c73613b6633b7101b`。单题结构、资产和 3 个 Schema 均为 0 个问题。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `automationbench4044-oracle-three-way-match-20260829`，trial 为 `automationbench-4044-three-way-m__WcfKJsw`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 6 页已逐页目视；来源对齐表、决策表、异常卡、原生金额图和 AP 通知页无重叠裁切，`visual_qa.py` 返回空问题列表。

## 视觉评分适配（visual-v1）

- 模型逐题读取并检查了任务指令、配置、确定性 Rubric、任务专属 Judge、负例预期、原适配说明、政策 EML、上下文 JSON，以及 Excel 三张工作表中全部四笔 Invoice/PO/Receiving 记录。
- Golden `Three_Way_Match_Exception_Briefing.pptx` 共 6 页，已先查看 3×2 接触表，再逐页查看 2000×1125 原图：批次结论、三源登记表、匹配决策表、两类异常处置、金额对比图、状态/通知/政策澄清。
- 本题允许多种专业财务简报布局和图表表达，采用 `reference_free`；Golden 仅用于任务理解、视觉锚点制定和校准，不作为候选必须像素复制的模板。
- `V001` 聚焦首屏批准/异常和金额概览，`V002` 聚焦两张高密度表的跨页追踪，`V003` 聚焦价格复核与数量冻结的并列区分，`V004` 聚焦四组金额柱及偏差解释，`V005` 聚焦最终状态和 AP 通知闭环，`V006` 聚焦财务控制视觉语言与六页节奏；这些标准避开确定性 Judge 已精确覆盖的金额、状态、来源和可编辑性。
- Golden 视觉风险：第 5 页 VI-001 的 5000/5100 与 VI-004 的 4400/4400 数据标签明显互相覆盖；真实视觉评分必须据此扣分，不得降低 Rubric 或修改 Golden 伪造满分。
- Golden 已用固定 Codex CLI 0.148.0、the configured VLM provider 连续真实评分 5 次，visual reward 为 `0.871875 / 0.881250 / 0.881250 / 0.881250 / 0.881250`，均值 `0.879375`、总体标准差 `0.003750`（小于 0.05）；确定性分均为 `1.0`，70/30 综合分均值为 `0.963813`。
- 任务特定中等负例只把第 6 页右侧 AP 通知面板整体左移 2.0 英寸，使其明显叠压四行 `APPROVE / REVIEW / HOLD` 决策说明，破坏状态序列到通知动作的扫描闭环；6 页仍可渲染、内容仍存在，确定性 Judge 仍为 `1.0`。真实 VLM visual reward 降至 `0.781250`，70/30 综合分为 `0.934375`，严格低于 Golden 均值且高于空白或损坏交付。
- 评分证据保存在 `.coworkevo/vlm-calibration/ppt/automationbench-4044-three-way-match-exception-briefing/`，不提交图片或模型响应。
