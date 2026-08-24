# 逐题改造记录：OmegaUse-OfficeVal officeval_086

## 为什么选择这道题

- 原始 Query、XLSX 数据表和 PNG 参考图均来自公开数据集，不是本项目自编的业务场景或输入数据。
- 它不只要求“有一张饼图”，而是要求主类别和“其他”明细同时成立、两图位置和连接线参考截图、最终必须原生可编辑，适合 `charting` 家族。
- 上游记录的人类完成时间为 65 分钟，且提供 2389 行的逐题 Verifier，可追溯并能评估原生图表结构。

## 原始来源

- 仓库：`baidu-frontier-research/OmegaUse-OfficeVal`
- 精确 commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- 上游任务：`data/tasks/officeval_086.json`
- 上游 Rubric：`data/rubrics/officeval_086.json`
- 上游 Verifier：`verifiers/officeval_086_verifier.py`
- 输入资产：Hugging Face `task_files/officeval_086/数据1.xlsx` 和 `image.png`
- 授权：Apache-2.0，上游 LICENSE 已保留。

## 实际检查过的输入

- `数据1.xlsx` 原始只有 Sheet1，共 7 行、3 列、无图表、无图片；主类别比例是 `0.55 / 0.34 / 0.11`，“其他”明细比例是 `0.15 / 0.65 / 0.20`。
- 原表使用“一级产品、二级产品、三级产品”和“优等品、合格品、残次品”作脱敏占位名，Query 明确给出它们分别应映射为“人员薪酬、物料、其他”和“交通、住宿、餐饮”。
- `image.png` 为 960×540 像素，展示一个较大左饼、一个较小右饼和两条虚线的视觉关系；图中的大洲与国家文字只是布局示意，不应被抄进最终业务图表。
- 两个输入的 SHA256 分别为 `20613e39c87d8dd68e27c3368b7014ac09dcec3a77d94e56d8548880aa246743` 和 `72786b96b5ccb28c265bb5d6dc355dd719e35add623fb596848c662eea01b62c`。

## CoworkEvoBench 的逐题改造决策

- 原始 Query 原文保留，只增加输入名、固定输出路径、Sheet1 数值保护和“不得用截图代替”边界。
- 参考图展示的是一个复合视觉，但不强制某个 Excel 专有的 `Pie of Pie` 对象。参考答案使用两个直接链接 Sheet1 的原生饼图和两条可编辑虚线，在 Excel、WPS 和 LibreOffice 之间更容易保持稳定，也符合原 Query 的“要可编辑”。
- 左图使用蓝/橙/青绿，右图使用紫/浅绿/蓝，扇区边界为 1 磅白线；主饼图和拆分饼图的可见直径比约为 `1.41`，水平间隔约 `3.10 cm`。
- Judge 在上游 Verifier 之外，另外检查 Sheet1 数值保留、工作表数量、两个图表的 Sheet1 数据引用、无嵌入图片代替、LibreOffice 实际渲染和统一 Judge schema。
- 参考解法是本题独立的数据映射和图表几何，不调用任何批量造题器。

## 已完成的逐题验证

- 参考解法从原始 XLSX 独立 replay 成功，Golden SHA256 为 `dc2149686a2b6e19f6431bf0c96abb8809dcc59f0b4b550cf0bcc01fde359a45`。
- 上游逐题 Verifier 的 8 个得分点全部通过，得分 `12/12`；本项目扩展 Judge 得分为 `1.0`。
- Golden 已由 LibreOffice 实际渲染为 2 页 A4 PDF，并人工查看 Sheet1 数据页与 Sheet2 整页图表；六个标签、百分比、两饼大小关系和两条虚线均可见。
- 空目录、损坏 XLSX 和未作答原文得分均为 `0`。
- 保留全部数据和两个原生图表，但移除虚线且把右饼放大到与左饼相同的中等缺陷样本得分 `0.730769`，低于 `0.75` 通过线。

## 尚未完成的发布门槛

- 需要由至少一个非参考 Agent Harness 独立作答，并用真实失败样本继续校准几何容差。
- 正式发布前需用 Microsoft Excel 或 WPS 再做一次兼容性抽检，特别检查虚线形状和数据标签位置。
