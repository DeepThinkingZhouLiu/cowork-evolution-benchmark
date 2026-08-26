# 逐题改造记录：OmegaUse-OfficeVal officeval_070

## 为什么选择这道题

- 原始任务来自经过脱敏和输入重建的真实办公请求，不是本项目编写的场景或 Query。
- 只有 1 页但含 65 个对象，六个正文板块同时存在框外溢出、编号标题挤压和跨区重叠，能真实考查局部版式修复与未修改内容保护。
- 上游提供原始 PPTX、Rubric 和 882 行逐题 Verifier，可追溯且能作为 Judge 基线。
- 上游记录的人类完成时间为 32 分钟，适合作为中等难度的 `deck_polishing` 候选。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_070.json`
- 上游 Rubric：`data/rubrics/officeval_070.json`
- 上游 Verifier：`verifiers/officeval_070_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_070/技术路线图.pptx`

## 实际检查过的输入

- PPTX 为 1 页竖版，尺寸约 20.00×28.29 厘米，共 65 个对象。
- 页面、所有文字和流程框已用 LibreOffice 实际渲染检查。
- 六个底框是背景 SVG 图片，编号、标题和正文是单独的可编辑文本框。
- 原始六个标题为 12.2 磅、正文为 8.6 磅，且多个文本框宽度大于底框；渲染结果中编号与标题粘连、正文越过框线，并压向相邻板块。
- 输入 SHA256 为 `07c278c83919f7e33d42b44324fef2fc8090b181be1b566b3b8c212002509875`。

## CoworkEvoBench 改造决策

- 原始 Query 原文保留，只增加输入、固定输出路径、离线和未修改内容保护边界。
- 只调整 18 个目标文本框，即 6 个编号、6 个标题和 6 个正文；不重绘背景，不替换其他对象，不改变任何文字事实。
- 以六个真实底框的几何位置为锚点，标题使用 11 磅加粗，正文使用 8 磅常规字重和 0.95 倍行距。
- 上游 Verifier 原样保留；本项目 Judge 补充其未实现的页面数/尺寸门槛、非目标对象保护、对象重叠、实际渲染和 schema 输出。
- 参考解法是这一题独立的几何修改，不共享批量生成器。

## 已完成的逐题验证

- 参考解法从原始 PPTX 独立 replay 成功，Golden SHA256 为 `3c16346c493f01d0f5f77cae024467e79a421b222251c21e648a4882583e693a`。
- 上游逐题 Verifier 得分为 `13/13`；本项目扩展 Judge 得分为 `1.0`。
- Golden 经 LibreOffice 实际渲染为 1 页 PDF，并人工查看整页。六组编号、标题和四行正文均完整位于对应底框，未压住箭头或相邻板块。
- 原始未作答文件得分为 `0.090909`，损坏文件和空目录得分为 `0`，均不能通过。
- 将第二板块正文故意向左移出底框的部分完成样本得分为 `0.545455`，位于原始未作答与 Golden 之间，证明 Judge 具有基础分数梯度。
- 正例及上述三类负例的 Judge 结果均通过 `schemas/judge-result.schema.json`。

## 尚未完成的发布门槛

- 需要由至少一个非参考 Agent Harness 独立作答并进行难度 Rollout。
- 需要从真实 Agent 输出中再选择自然产生的中等缺陷样本，避免只依赖人工位移样本校准。
- 正式发布前仍需用 PowerPoint 或 WPS 再做一次兼容性抽检，确认 SVG 背景和字体在目标环境下显示一致。
