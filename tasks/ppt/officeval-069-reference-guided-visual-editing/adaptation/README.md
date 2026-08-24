# 逐题改造记录：OmegaUse-OfficeVal officeval_069

## 为什么选择这道题

- 原始任务来自经过脱敏和输入重建的真实办公请求，Query、PPTX 和参考截图都不是本项目编写或合成的。
- 参考截图通过不同颜色的标注框和箭头指定“哪个对象要向哪个范例对齐”，能真实考查多模态参考理解、局部样式修复和未修改内容保护。
- 上游记录人类完成时间为 38 分钟，并提供逐框 Rubric 与 Verifier，适合作为 `visual_editing` 真实题。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_069.json`
- 上游 Rubric：`data/rubrics/officeval_069.json`
- 上游 Verifier：`verifiers/officeval_069_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_069/`

## 实际检查过的输入

- PPTX 为 1 页 50.80×28.575 厘米的横版流程图，共 64 个对象，包含 62 个可编辑形状和 2 条可编辑连接线，不是截图。
- 参考 PNG 为 960×546，红、粉、绿标注框与同色箭头一共指向 10 个需要修改的圆角矩形。
- 逐对象检查后确认，时间轴框应对齐浅蓝范例；公司治理类框应对齐 4.45×2.08 厘米的浅蓝范例；运营与监督类框应对齐同尺寸的浅粉范例。
- 输入 PPTX SHA256 为 `2eee79777713766520578b96ed0577216cc2b323facb2d96f99b3f6b0a64caac`，参考 PNG SHA256 为 `097b870d2197d95dc4a0203156c4494cf3519d00703fba2cdf0d8bec510198a2`。

## CoworkEvoBench 改造决策

- 原始 Query 逐字保留，只增加输入文件名、固定输出路径、离线边界和未标注对象保护要求。
- 参考结果是在真实原 PPTX 上逐框修改，不是重画整页，也没有调用批量生成脚本。
- 两个较长的粉色框仅重排原文换行，没有增删或改写内容。
- 上游 Verifier 原样保留；本项目 Judge 补充页面尺寸、64 对象可编辑性、未标注对象保护、防截图覆盖、实际渲染和统一结果 schema。
- `solution/solve.py` 是这道题独立的 replay 记录，只包含这 10 个真实目标的人工映射，不是可以换数据批量生成题目的工具。

## 已完成的逐题验证

- 参考解法可从原始 PPTX 独立 replay，Golden SHA256 为 `efeb17cf0eb0bbdbc370755e1303e259aa06c091048bf14eba97b6128a80c639`。
- 上游逐题 Verifier 的 10 条规则全部命中，得分 `30/30`。
- 本项目扩展 Judge 对 Golden 得分为 `1.0`。原始未作答文件得分为 `0.526316`；只完成四个标注框的部分作答得分为 `0.736842`，低于 `0.75` 通过线。
- 在原稿上叠加整页截图的错误交付触发可编辑性门槛，得分为 `0`；损坏文件和空目录也都得分 `0`。
- Golden 及上述四类负例的 Judge 结果均通过 `schemas/judge-result.schema.json`。
- Golden 经 LibreOffice 实际渲染为 1 页 PDF，并已人工查看整页；10 个目标框文字可读，其他流程图对象保持可编辑。

## 尚未完成的发布门槛

- 需要由至少一个非参考 Agent Harness 独立作答并进行难度 Rollout。
- 正式发布前仍建议用 Microsoft PowerPoint 或 WPS 做一次字体与连接线兼容性抽检。
