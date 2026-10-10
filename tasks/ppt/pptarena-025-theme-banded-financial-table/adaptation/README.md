# PPTArena Case 25 逐题适配记录

## 适配契约

- 角色：财务演示文稿编辑。
- 输入：1 页 16:9 PPTX，含 3 个页面对象、4 个媒体部件和 1 个 5×4 原生财务表格。
- 真实意图：使用文稿现有主题色，把与深绿背景冲突的亮黄表格改成深色表头、高对比文字和 A-B-A-B 交替数据行。
- 主交付物：`FixTableThemeColor_GroundTruthA.pptx`。
- 不可妥协约束：表格内容、行列、合并状态、位置、尺寸、边距与对齐保持；标题、页码、背景、主题、母版、版式、备注和媒体保持；表格必须原生可编辑。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 25: Match Table Theme and Colors。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 和 Hugging Face 数据卡均标记 MIT；许可与数据说明副本保存在本目录。
- 上游 Query、style target、路径、类别、输入审计和适配边界逐字段保存在 `upstream-case.json`。

文稿使用了外部模板图形素材。PPTArena 数据集的 MIT 标记不会自动消除第三方模板权利风险，正式公开 Release 前必须单独复核或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/FixTableThemeColor_TestA.pptx` | 558750 | `d0b3c9d5b9988a375252315d3300b3237f25684af23ed77a4330aa8021d10c30` | 1 页原尺寸渲染；3 个对象、4 个媒体、49 个包成员、表格文字、几何、填充、对齐和主题色审计 |
| `solution/Golden/FixTableThemeColor_GroundTruthA.pptx` | 555710 | `fa3fbef6d9d97a8a5f71f46dff48ef0399555050ecddb6555ba3bd7face53acb` | 从冻结输入逐单元格重建；原尺寸目视、OOXML 差异、媒体哈希和 LibreOffice 打开渲染审计 |
| `adaptation/upstream-ground-truth-media-mutated.pptx` | 599479 | `f667946a09e9740f85181075ef0fb8c4f1c59d05dec0d562443c04a278c95ddd` | 上游人工 Ground Truth；表格样式可用，但确认无关改写 2 个背景媒体和多个文档状态部件 |
| `adaptation/build_golden.py` | 4437 | `8c0a96aa668b1dd5e604b9627e31f2261d58b69d2c6cdda373778eb04c792831` | 任务专属重建脚本；冻结输入哈希、5×4 文字和逐行颜色约束被显式写入 |
| `adaptation/upstream-case.json` | 3153 | `e07968efe97c97e5febce1d0ed358e047b373c348aee3865e9810dcc3c5a8039` | 上游 Query、style target、来源版本、输入清单与不沿用的上游副作用逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入与上游 Ground Truth 审计

冻结输入为 960×540pt 的单页文稿，标题为 `Financials`，页码为 `1`。唯一原生表格的位置为 x=66pt、y=143.75pt，尺寸为 828×319.23pt，5 行×4 列；表头为 Quarter、Revenue ($)、Expenses ($)、Net profit ($)，Q1–Q4 的 12 个金额均已逐格核对。输入的 20 个单元格都是 `#FFFF00` 亮黄填充，与页面深绿背景明显冲突。

上游 Ground Truth 的表头使用 `accent6` 的 50% 亮度，数据行使用 `tx2` 的两档亮度变体，全部文字为 `bg1`；这一表格方案符合上游 style target。但与输入相比，它共改动 8 个包部件：`docProps/app.xml`、`docProps/core.xml`、`docProps/thumbnail.jpeg`、`ppt/media/image1.png`、`ppt/media/image3.png`、`ppt/presentation.xml`、`ppt/slides/slide1.xml` 和 `ppt/viewProps.xml`。两张 PNG 解码后的 RGBA/透明度均发生真实像素变化，不是单纯压缩重写；虽然当前 LibreOffice 渲染的表格区外画面没有可见差异，但它仍违反“只改表格”和媒体字节保持契约，因此仅作设计证据和专项负例。

## 本地 Golden 构造

本地 Golden 由制题者在完整读取财务数据、表格几何、主题色和上游副作用后逐单元格重建，不是通用批量 mock：

- 从冻结输入出发，49 个包成员中只改 `ppt/slides/slide1.xml`。
- 表头 4 格使用 `accent6 + lumMod 50000`，文字使用 `bg1`。
- Q1/Q3 行使用 `tx2 + lumMod 75000 + lumOff 25000`，Q2/Q4 行使用 `tx2 + lumMod 50000 + lumOff 50000`，文字使用 `bg1`。
- 20 格文字、表格框、4 列宽度、5 行高度、单元格边距、对齐、标题、页码和备注不变。
- 主题、母版、版式、4 个媒体和其他 48 个包部件字节保持。

## Judge 设计

本地 Judge 不读取 Golden，也不比较 Golden 哈希；它只从候选 PPTX 和冻结输入独立重算。

- 使用 `python-pptx` 检查唯一 5×4 原生表格，逐格核对文字、合并状态、边距、对齐、位置、尺寸、行高和列宽。
- 解析冻结主题的 `dk2/tx2/accent1–accent6`，按 OOXML HLS 亮度变换解析 `lumMod/lumOff/tint/shade`；接受不同的主题色及其明暗变体，拒绝泛化黑白灰或与原主题无关的颜色。
- LibreOffice 实际渲染后，从每个单元格内部采样可见填充；表头必须行内同色、深色且与浅色文字对比不低于 4.5。
- 4 个数据行必须行内同色，A 行对 A 行、B 行对 B 行的颜色距离不超过 14，A/B 之间颜色距离不低于 28，文字对比不低于 2.55；该阈值会接受经目视合格的上游风格。
- 与亮黄 `#FFFF00` 的颜色距离必须不低于 90，用于拒绝任意单元格残留旧填充。
- 比对非表格对象的规范化签名、备注、版式、主题、母版、媒体哈希 Counter、关系和包成员集，专门拒绝上游 Ground Truth 的非目标媒体改写。
- 遮罩表格旧新边界后比较表外像素，相似度必须不低于 0.9985；表格可读性由渲染填充和 OOXML 字色联合检查。
- 规范化掉合法的单元格填充、字色和表格样式后的内容保持只占 1 分，不会将答案锁死到某一个 Golden 字节流。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地逐单元格重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立 `accent3` 蓝色表头与蓝色行带答案 | 1.000000 | 通过 | 颜色、字色与 Golden 不同，但满足主题、对比度和 A-B-A-B 契约 |
| 独立 `dk2 + accent6` 深绿/浅青答案 | 1.000000 | 通过 | 使用另一组主题色、浅色数据行与黑字，仍完全合法 |
| fresh replay | 1.000000 | 通过 | 从空输出目录回放，产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.990000 | 中等 | 功能、原生表格和渲染全部正确，只有 1 分包骨架字节保持失败 |
| 空交付 / 损坏 PPTX / 未修改原件 | 0.000000 | 拒绝 | 缺文件、无法打开或表头/行带未完成 |
| 上游 Ground Truth | 0.000000 | 拒绝 | 虽表格样式正确，但 2 个背景媒体被无关改写 |
| 表头过亮 / 非主题表头 / 表头深色字 | 0.000000 | 拒绝 | 表头明暗、主题相关性或高对比失败 |
| 无行带 / A-A-B-B / 行内异色 / 行带过于接近 | 0.000000 | 拒绝 | A-B-A-B 结构或可辨性失败 |
| 亮黄残留 / 数据字与背景同色 / 泛化黑白灰 | 0.000000 | 拒绝 | 旧填充、文字对比度或原主题相关性失败 |
| 改季度 / 移动或缩放表格 / 改行高列宽 | 0.000000 | 拒绝 | 内容或几何保持失败 |
| 改标题 / 改页码 / 改媒体 / 改主题 | 0.000000 | 拒绝 | 非目标页面内容或文稿资产保持失败 |
| 删表格 / 表格截图压平 / 额外对象 / 额外页 | 0.000000 | 拒绝 | 原生可编辑性、对象数或页数失败 |
| 答案路径标记 | 0.000000 | 拒绝 | 交付卫生、对象保持和反泄漏检查失败 |

`expectations.json` 共记录 33 组 Golden、独立正例、fresh replay、中等样例和负例，实际结果全部与期望一致。同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena025-oracle-theme-banded-financial-table-20260828`，trial 为 `pptarena-025-theme-banded-financ__dGbUfFU`，0 个异常；容器产物 SHA256 为 `fa3fbef6d9d97a8a5f71f46dff48ef0399555050ecddb6555ba3bd7face53acb`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-025-theme-banded-financial-table
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-025-theme-banded-financial-table
    harbor run --path tasks/ppt/pptarena-025-theme-banded-financial-table --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --jobs-dir /tmp/cowork-harbor-jobs --job-name pptarena025-oracle-theme-banded-financial-table-20260828

## 视觉复核

- 已按 presentation-skill 流程渲染并原尺寸查看冻结输入、上游 Ground Truth 和本地 Golden 的全部 1 页。
- 已原尺寸查看两个独立合法变体；蓝色和深绿/浅青两组方案均与原深绿页面协调，行带清晰，数字无裁切、遮挡或错位。
- `visual_qa.py` 对冻结输入、上游 Ground Truth、本地 Golden 和两个合法变体都只报告原文稿已存在的标题文本框过高启发式警告；该框和标题未被本题修改，不构成新的视觉缺陷。
- 本地 Golden 的表格区外渲染相似度为 0.998893，高于 0.9985 保持阈值；剩余微小差异来自 JPEG 反锁和表格边缘抗锯齿。

## 剩余风险

- 正式公开 Release 前必须复核模板背景图形的第三方权利；如无明确授权，需替换为可公开素材并保持相同的表格换色语义。
- LibreOffice/PDF 渲染可检查行带、对比度和布局，但不能穷尽 Microsoft PowerPoint 对内置表格样式的全部继承规则；正式 Release 前建议在 PowerPoint 中人工点选单元格、查看填充主题关系并编辑一个金额后撤销一次。
- 当前 L2 / Medium 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 任务专属视觉评分补充

制题模型重新完整读取了 instruction、task.toml、deterministic rubric、完整 judge、expectations、solution、本地 Golden 重建脚本、upstream case 和全部适配记录，并以原尺寸渲染查看了冻结输入、本地 Golden 与上游媒体被改写的 Ground Truth。视觉评分采用 `reference_free` 模式：任务允许从原主题中选择不同的深色表头和交替行带组合，Golden 只用于理解合格样式，不应把有效候选锁死为一组色值。

`visual_rubric.json` 为本题单独定义五个维度：财务表格与深绿页面的主题协调、统一深色表头及高对比浅字、Q1–Q4 严格 A-B-A-B 行带、20 格季度/金额的可读对齐，以及表外页面保持与无截图/遮罩痕迹。每项都写明 0、0.5、1 的可见锚点和原尺寸检查证据。

Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、the configured VLM provider 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。

首轮负例把 Q4 错设为与 Q3 相近的深绿色，但 the VLM evaluator 仍误读为正常 A-B-A-B 并给出 1.0，因此拒绝该负例。合格 v2 负例保留 5×4 数据、统一深色表头、居中对齐、原生表格、深绿页面和页码，但将 Q4 整行改为可读的暗酒红色，使行带明确变为 A-B-A-C。其代码分为 0.0，the VLM evaluator 视觉分为 0.890000，按 70/30 权重得到综合分 0.267000；视觉分严格低于 Golden，且高于空白或损坏文件。本题 7 次 the VLM evaluator 调用均一次成功，未发生基础设施失败；任何后续 API、渲染或解析失败只记为评测基础设施失败，不计作 0 分。
