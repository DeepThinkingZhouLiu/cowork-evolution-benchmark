# PPTArena Case 59 逐题适配记录

## 适配契约

- 角色：教育传播演示文稿编辑。
- 输入：1 页户外活动主题演示文稿，共 15 个页面对象和 3 个图片媒体。
- 真实意图：只将目标句中的 Berkeley 和 Illinois 两个 URL 分别设为指向自身的可点击外部超链接。
- 主交付物：`ConvertHyperLink_GroundTruth.pptx`。
- 不可妥协约束：两个 URL 的全部字符都正确链接；周围文字和其他对象不带链接/动作；原页面内容、布局、图片、样式和顺序保持。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 59: Convert Hyper Link。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入和 Ground Truth 路径已逐字段保存在 `upstream-case.json`，未将其伪称为本项目自编数据。

文稿包含上游插画图片。PPTArena 数据集的 MIT 标记不会自动消除图片本身的第三方权利风险，正式公开 Release 前必须单独复核授权或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/ConvertHyperLink_Test.pptx` | 456747 | `33132900f1422268aec0a20ba9e6ba0c2be1e6b9a4baeb9f0887c371506792f5` | 原尺寸渲染；所有 15 个对象、文字/run、几何、图片/裁剪、关系和包部件审计 |
| `solution/Golden/ConvertHyperLink_GroundTruth.pptx` | 454306 | `2e976bad97a869c58f0fe380bd04911a62a613c5d271c446e7e8cb35cd51a94e` | 本地人工重建；超链接目标、关系、OOXML 差异和原尺寸视觉检查 |
| `adaptation/upstream-ground-truth-unfit.pptx` | 456636 | `8438f624ef58a86a9062e1e50e6d6c1ba12d03a95579795cf0dd4af6d3ba6d0f` | 上游人工 Ground Truth；确认链接正确，但整页背景大图被移动且渲染严重裁切 |
| `adaptation/upstream-case.json` | 1121 | `dcf2b0b9c426151fd8c96deaf3b17e232cda1b2aaa0905210bb63097dd807c39` | Query、style target、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入和上游 Ground Truth 审计

冻结输入为 1 页、13.333×7.5 英寸，共 15 个页面对象、3 个 PNG 媒体和 45 个 ZIP 包成员。目标文本框形状 ID 为 918，位置约为 `x=0.75, y=4.57, w=9.336, h=0.283` 英寸；原始句子是单一 12pt run，无任何 hyperlink 关系。

上游 Ground Truth 把该句拆成 5 个 run，两个 URL 分别指向正确的 `TargetMode=External` 关系，链接语义本身正确。但它同时将整页背景图的几何从约 `x=-3.900, y=-7.050` 改为 `x=-3.982, y=-7.411` 英寸，并额外修改 `presentation.xml`、`viewProps.xml` 和缩略图等部件。LibreOffice 原尺寸渲染中，左上 Logo/导航与右侧背包均被明显裁切，违反“其他内容不变”，因此只作为证据和负例。

## 本地 Golden 构造

本地 Golden 由制题者在完整读取原句、run 样式、关系和上游意图后逐题重建，不是通用批量 mock：

- 从冻结输入出发，只改写 `ppt/slides/slide1.xml` 和 `ppt/slides/_rels/slide1.xml.rels`。
- 将原单 run 拆成 `Visit `、Berkeley URL、` and `、Illinois URL、` for more details` 五段，复制原 12pt 运行属性。
- 只向两个 URL run 写入 `a:hlinkClick`，并分别绑定到完全匹配的 HTTPS 外部目标。
- 其他 43 个包成员、3 个媒体、所有对象几何、裁剪、文字、版式和备注字节保持。
- 原尺寸视觉检查显示，除 URL 采用演示软件默认 hyperlink 样式外，页面其余区域与冻结输入逐像素一致。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，并根据任务契约独立重算。

- 将目标句每个字符映射到所在 run 的 hyperlink 目标；两个 URL 的全字符覆盖必须精确，周围字符必须全部为空。
- 允许一个 URL 被拆成多个 run，但每段都必须指向同一正确目标；这避免锁死 Golden 的 run 数和关系 ID。
- 扫描整个 PPTX 的 `.rels`，必须恰好有两个 slide1 外部 hyperlink 关系，目标 Counter 与两个 HTTPS URL 完全相等。
- 扫描整页 `hlinkClick/hlinkMouseOver`，拒绝整句链接、对象级链接、额外动作、漏链、交换链接和 HTTP/内部目标。
- 目标句必须逐字保持，每个 run 保持 12pt 和原字体/强调属性；只有 URL 允许默认颜色或下划线。
- 对比除目标 run 分段外的 15 个对象语义签名、几何、顺序、图片哈希/裁剪、版式和备注。
- 媒体哈希 Counter、包成员集和对象数独立检查，拒绝截图压平、替换图片和额外素材。
- LibreOffice 实际渲染冻结输入与候选文稿；对 URL 文本框扩展区域做掉码后，页面其余像素相似度必须不低于 0.9985。
- 规范化掉合法 hyperlink 后的内容 XML、关系和关键包部件字节级保持只是 2 分正向分项，不会把合法答案锁死为某一个 Golden 字节流。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地人工重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立显式蓝色下划线答案 | 1.000000 | 通过 | 使用不同关系 ID 和显式样式，两 URL 目标正确 |
| 独立 URL 多 run 答案 | 1.000000 | 通过 | 每个 URL 分两段，但逐字链接覆盖正确 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.980000 | 中等 | 功能、对象和渲染正确，仅 2 分字节级保持分项失败 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺文件、无法打开或无功能链接 |
| 上游错位 Ground Truth | 0.000000 | 拒绝 | 背景大图几何和页面渲染保持失败 |
| 只改样式 / 只链一个 / 交换目标 | 0.000000 | 拒绝 | 缺真正 hyperlink、覆盖不完整或 URL 错链 |
| 整句链接 / 周围文字链接 / 对象级链接 | 0.000000 | 拒绝 | 链接范围越界 |
| HTTP 目标 / 缺 External 属性 | 0.000000 | 拒绝 | 外部 HTTPS 关系不精确 |
| 改字号 / 改 URL 文字 / 改标题 | 0.000000 | 拒绝 | 原字号或可见内容保持失败 |
| 移背景 / 改图片 / 答案标记 / 额外页 | 0.000000 | 拒绝 | 布局、媒体、交付卫生或页数失败 |

包含 fresh replay 在内共 25 组正例、边界例和负例，全部与 `expectations.json` 一致。同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena059-oracle-two-url-links-20260828`，trial 为 `pptarena-059-two-url-clickable-h__VXnLZVR`，0 个异常；容器产物 SHA256 为 `2e976bad97a869c58f0fe380bd04911a62a613c5d271c446e7e8cb35cd51a94e`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-059-two-url-clickable-hyperlinks
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-059-two-url-clickable-hyperlinks
    harbor run --path tasks/ppt/pptarena-059-two-url-clickable-hyperlinks --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 流程渲染冻结输入、上游 Ground Truth、本地 Golden 和显式蓝色下划线合法变体。
- 已原尺寸查看上述页面；本地 Golden/合法变体的 Logo、导航、主标题、背景插画和背包完整，上游 Ground Truth 的左右裁切可稳定复现。
- `visual_qa.py` 对冻结输入和本地 Golden 都报告 0 条布局警告。
- 本地 Golden 的 URL 文本区外与冻结输入渲染像素相似度为 1.0。

## 剩余风险

- 正式公开 Release 前必须复核上游插画图片的第三方权利；如无明确授权，需替换为可公开素材并保留同等超链接任务语义。
- LibreOffice PDF 渲染能检查布局和 hyperlink 默认样式，但无法代替真实鼠标点击；可点击性已由 OOXML `hlinkClick` 与 External relationship 交叉检查。正式 Release 前仍建议用 Microsoft PowerPoint 点击两个 URL 各一次。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本适配记录、上游 Case 与元数据、本题 `build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入、本地 Golden 和上游 Ground Truth；核对了目标句的五段语义、两个 URL、15 个页面对象、3 个媒体、外部链接关系以及只允许微编辑的保持边界。
- 冻结输入、本地 Golden 和上游 Ground Truth 均已原尺寸渲染并逐页查看：Golden 仅让 Berkeley 与 Illinois 两个 URL 呈现默认链接样式，周围文字仍是普通 12pt 正文；大写主标题、浅蓝地图背景、左上 Logo、顶部导航、右上装饰线和右侧背包保持完整。上游 Ground Truth 虽然链接正确，但背景大图几何发生变化，淡色地图/云纹轮廓相对标题和页面边缘清楚错位。
- 本题是对既有单页的精确微编辑与严格版式保持，因此使用 `reference`。专属视觉 Rubric 检查两个 URL 的视觉链接范围、完整句子的单行排版、地图/Logo/导航/背包的背景保持、主标题—插画—链接句的信息层级以及页面是否只留下必要改动。
- URL 真正可点击、逐字符目标覆盖、HTTPS External 关系、对象数/几何、媒体哈希、包成员和可编辑性继续由确定性 Judge 负责；VLM 只评估视觉链接边界、排版、背景构图保持和微编辑完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、the configured VLM provider 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次确定性分与综合分也均为 `1.0000`。
- 先后实测了上游背景错位答案和将背景形状横移 0.5 英寸的弱负例；虽然两者确定性分均为 `0.0000`，但 the VLM evaluator 都给出视觉分 `1.0000`，说明当前 VLM 对淡色地图纹样的轻微位置变化不敏感，因此它们只保留为校准失败记录，不冒充有效中等负例。
- 最终中等可见负例从 Golden 出发，仅把 `Visit`、`and` 和 `for more details` 三段非 URL 文字改为高饱和洋红色并加下划线；两个真实 URL、完整句子、单行位置、标题、背景和背包仍保持，文稿可正常打开。已原尺寸渲染并亲眼确认该缺陷直接对应 V001/V004/V005；其确定性分 `0.0000`、视觉分 `0.7690`、综合分 `0.2307`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 8 次 VLM 调用（5 次 Golden、2 次弱负例诊断、1 次有效负例）均成功完成，没有基础设施失败或被错误计作 0 分的样本。
