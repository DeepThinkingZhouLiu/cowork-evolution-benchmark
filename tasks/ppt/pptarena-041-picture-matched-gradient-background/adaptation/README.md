# PPTArena Case 41 逐题适配记录

## 适配契约

- 角色：医学图表演示文稿编辑。
- 输入：1 页《Overall interactions heat map》演示文稿，中央是一张内嵌 TIFF 热力图。
- 真实意图：从图片的浅色像素或文稿 `accent1` 主题色构造上近白、下浅蓝灰的纵向渐变背景。
- 主交付物：`MatchBackgroundtoPicture_GroundTruthA.pptx`。
- 不可妥协约束：必须使用原生幻灯片渐变背景；图片字节、裁剪、几何和渲染内容保持；标题、页脚、字体、链接、版式和对象顺序保持。
- 可接受变体：图片可选增加不超过 2pt 的蓝色或 `accent1` 细边框。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 41: Match Background to Picture。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入与 Ground Truth 路径已保存在 `upstream-case.json`，未将其伪称为本项目自编数据。

文稿内部嵌入已发表医学论文的热力图与 DOI。PPTArena 的 MIT 标记不会自动消除该第三方图表的权利风险，正式公开 Release 前必须单独复核授权或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/MatchBackgroundtoPicture_TestA.pptx` | 149374 | `8c789d83d7063549ab3a198a7363e7e6d9feda84ea20865ae00a6307d3be52cd` | 原尺寸渲染；文字、对象、几何、图片字节/裁剪、主题色和包部件审计 |
| `solution/Golden/MatchBackgroundtoPicture_GroundTruthA.pptx` | 149483 | `ac8d5b8457fa8be04d58acd326201e66a21caf505210cc90eb82aa92fddae26d` | 本地人工重建；原尺寸视觉检查、OOXML 差异和中央图片渲染对照 |
| `adaptation/upstream-ground-truth-unfit.pptx` | 178017 | `4ec666fc004cc5dd06003bd9e11866c215aacc3da2a6cc2fd5252092777bac47` | 上游人工 Ground Truth；结构和原尺寸渲染检查，确认中央 TIFF 严重错渲染 |
| `adaptation/upstream-case.json` | 1225 | `a42d4d8c462e9a8f2ae6e731aef2634b6aa4e772f556c94d579aa208e9f57925` | Query、style target、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入和上游 Ground Truth 审计

冻结输入为 1 页、10×7.5 英寸，包含 3 个页面对象：标题、DOI/引用页脚和中央 TIFF 热力图。热力图是 971×696 像素、146212 字节，页面几何为 `x=1.5122, y=1.25, w=6.9756, h=5.0` 英寸，四侧裁剪均为 0。图中常见的浅色包括近白、`(224,232,240)` 和 `(176,192,216)`；文稿主题 `accent1` 是 `#4472C4`。

上游 Ground Truth 使用 4 个 `accent1` 亮度变体的纵向渐变，并为图片增加蓝色细边框，这两点符合文字意图。但是原尺寸 LibreOffice 渲染显示，它将原本完整的中央 TIFF 热力图渲染为狭窄碎片/大片空白，违反“图片保持不变”。因此它只作为可追溯证据和负例，不作为本题 Golden。

## 本地 Golden 构造

本地 Golden 由制题者在读取原图、主题色和上游意图后逐页重建，不是用通用生成脚本批量 mock：

- 从冻结输入出发，只向 `ppt/slides/slide1.xml` 的 `p:cSld` 写入原生 `a:gradFill`。
- 顶部色标为 `accent1 + lumMod=5000 + lumOff=95000`，解析颜色约为 `(246,248,252)`。
- 底部色标为 `accent1 + lumMod=30000 + lumOff=70000`，解析颜色约为 `(199,213,237)`。
- 渐变角度为 90°，未添加可选边框，避免重现上游错渲染。
- 中央图片媒体字节、裁剪、几何、文字、链接、关系和其他关键包部件均与冻结输入保持。

原尺寸渲染中，背景侧带从顶部约 `(241,244,253)`、中部 `(223,229,243)` 过渡到底部 `(203,215,239)`；中央图片内部与冻结输入渲染保持。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，并根据任务契约独立重算。

- 解析第 1 页 OOXML，必须恰有一个幻灯片原生线性渐变背景，至少有两端色标，且角度为纵向。
- 接受两类合法颜色：`accent1` 亮度变体，或直接来自原图浅色像素的 RGB 端点。
- 对顶部近白、底部明亮浅蓝灰、上浅下深、两端色差和原图调色板距离设置可解释的数值边界。
- 对比标题、页脚、run 样式、超链接、对象顺序、几何、旋转、版式、图片哈希和裁剪。
- 可选边框必须不超过 2pt，且是蓝色或 `accent1`；粗红边框明确拒绝。
- 媒体哈希 Counter、包成员集、3 个原生对象和整页图片覆盖独立检查，防止丢媒体、截图压平和图形伪背景。
- LibreOffice 实际渲染冻结输入与候选文稿，中央图片内部像素相似度必须不低于 0.9985；同时独立检查背景侧带的上中下亮度和色相。
- 为允许合法细边框，像素比较避开图片边界的 LibreOffice 抗锯齿区域；图片字节、裁剪和几何仍由结构规则严格保护。
- 去除合法背景/边框后的内容 XML 和关键包部件字节级保持只是 2 分正向分项，不会把合法答案锁死为某一个 Golden 字节流。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地人工重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立 RGB 取样渐变 | 1.000000 | 通过 | 从冻结输入制作，不复制 Golden 的主题色参数 |
| 合法 1pt `accent1` 细边框 | 1.000000 | 通过 | 任务书允许的可选变体不被误拒 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.980000 | 中等 | 任务语义、对象和渲染正确，仅 2 分字节级保持分项失败 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺文件、无法打开或缺失渐变 |
| 上游错渲染 Ground Truth | 0.000000 | 拒绝 | 图片对象/实际渲染保持失败 |
| 纯色 / 图形伪背景 / 水平渐变 | 0.000000 | 拒绝 | 不是符合约束的原生纵向渐变 |
| 纯白 / 过暗 / 黄色 / 上深下浅渐变 | 0.000000 | 拒绝 | 亮度、蓝灰色系或方向不符合意图 |
| 改标题 / 移图 / 改裁剪 / 删图 | 0.000000 | 拒绝 | 原内容、几何、裁剪或媒体保持失败 |
| 粗红边框 / 答案路径标记 / 额外第 2 页 | 0.000000 | 拒绝 | 边框越界、未授权内容或页数变化 |

包含 fresh replay 在内共 23 组正例、边界例和负例，全部与 `expectations.json` 一致。同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena041-oracle-picture-gradient-20260828`，trial 为 `pptarena-041-picture-matched-gra__PqRqoCb`，0 个异常；容器产物 SHA256 为 `ac8d5b8457fa8be04d58acd326201e66a21caf505210cc90eb82aa92fddae26d`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-041-picture-matched-gradient-background
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-041-picture-matched-gradient-background
    harbor run --path tasks/ppt/pptarena-041-picture-matched-gradient-background --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 的流程渲染冻结输入、上游 Ground Truth 和本地重建 Golden。
- 已原尺寸查看三份文稿的第 1 页；冻结输入和本地 Golden 的完整热力图正常，上游 Ground Truth 的 TIFF 错渲染可复现。
- 已检查本地 Golden 与冻结输入的中央图片内部像素、背景上中下端点和独立 RGB 合法变体；渐变柔和、不遮挡标题/页脚，且与热力图蓝灰色阶协调。

## 剩余风险

- 正式公开 Release 前必须复核已发表医学热力图的第三方权利；如无明确授权，需替换为可公开素材并重做颜色真值。
- 当前已通过 OOXML、python-pptx 语义对象、LibreOffice 渲染、人工对照和 Harbor Docker 回放；正式 Release 前仍建议用 Microsoft PowerPoint 打开一次，确认 TIFF 与主题渐变在目标应用中一致。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、任务专属 `build_golden.py` / `solve.py`、冻结输入、上游不适用 Ground Truth 与本地 Golden，核对了渐变色标/方向、热力图媒体哈希/几何/裁切、主题 accent1、标题、引用与 DOI 链接。
- 冻结输入、本地 Golden 和上游 Ground Truth 均已原尺寸渲染并亲眼对照：输入为纯白底；本地 Golden 从近白顶部均匀过渡到明亮浅蓝灰底部，中央 2×4 热力图、8 个数值、行列标签、色标、标题与 DOI 均完整清晰；同时复核了上游文件的图片边界/兼容性风险。
- 任务允许主题亮度变体或从图片浅色区域取样的多种合法端点，因此使用 `reference_free`，不强迫候选复制 Golden 的颜色数值或是否加边框。专属视觉 Rubric 分别检查上浅下深的全页渐变、与热力图色谱的协调、热力图 8 数值/标签/色标的完整可读、标题—图—页脚层级与原生背景/边框的交付干净度。
- 原生 `a:gradFill`、纵向角度、色标结构/颜色边界、图片字节/几何/裁切、可选边框宽度、对象签名、媒体哈希、链接与原生可编辑性继续由确定性 Judge 负责；VLM 只补充渐变平滑度/柔和度、色调协调、图表真实可读性、层级和专业完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次综合分均为 `1.0000`。
- 有效中等可见负例从 Golden 手工定向改造：保留原生纵向渐变、近白顶部、中央 TIFF 字节/几何/裁切、3 个原生对象、标题和 DOI，但把底部 accent1 改为无亮度偏移的 `70%` 海军蓝，使底部过暗、与热力图的浅蓝灰色标竞争，并明显降低页脚引用/DOI 对比。已原尺寸渲染并亲眼确认：文稿仍可打开、热力图与主标题仍完整可读，但达到肉眼可见的中等配色/对比缺陷。该负例确定性分 `0.0000`、视觉分 `0.6900`、综合分 `0.2070`，严格低于 Golden 均值且高于空白/损坏交付。
- 首版负例只把底色调深到约 `(134,163,217)`，确定性规则因超边界给出 0 分，但视觉模型仍给出 `1.0000`，因区分度不足而未作为有效负例；该真实记录保留在忽略目录供审计。本轮无未解决基础设施失败，未将任何失败冒充为 0 分。
