# PPTArena Case 16 逐题适配记录

## 适配契约

- 角色：视觉传播演示文稿编辑。
- 输入：2 页横版 PPTX，共16 个原生页面对象和 4 个媒体部件。
- 真实意图：在第 2 页识别瀑布、螺旋楼梯和国际象照片，将它们正确对应 WATERFALL、STAIRCASE 和 CHESS 标签，整理成顶边对齐、等大近方形、不遮挡标签的可编辑三联画。
- 主交付物：`OrderPhotosGroundTruth.pptx`。
- 不可妥协约束：三图顺序与标签匹配；无旋转、等大、顶边对齐；标签完整可见；第 2 页显式白色纯色背景；封面、主题、青蓝装饰、粉色面板和标签不变。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 16: Curate Multi-Panel Photo Layout。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入、Ground Truth 路径和上游已承认的无关换肤副作用已逐字段保存在 `upstream-case.json`。

文稿含外部模板图、瀑布、螺旋楼梯和国际象摄影图片。PPTArena 数据集的 MIT 标记不会自动消除第三方素材权利风险，正式公开 Release 前必须单独复核授权或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/OrderPhotosTest.pptx` | 5676140 | `c17f8b37992c462006bf62a151b3f346df2dade7253f97027f69d715ab7f657c` | 2 页原尺寸渲染；16 个对象、4 个媒体、图片内容/裁剪、几何、旋转、文字、顺序、背景和包部件审计 |
| `solution/Golden/OrderPhotosGroundTruth.pptx` | 5673705 | `5791250ffc46eaa7a95b388d86c246b09030ce476b430b84a170d7c496c97f05` | 本地逐题重建；2 页原尺寸目视、OOXML 差异、媒体哈希和交付可读性审计 |
| `adaptation/upstream-ground-truth-unfit.pptx` | 5674936 | `661612919647915ea1c8ce6a4c71f4e9cb5857a369dbd7619205b2f93bc29fd0` | 上游人工 Ground Truth；确认无关换主题、改封面、标题不可见和标签遮挡问题 |
| `adaptation/build_golden.py` | 4004 | `887b160e3b79dd9e80d03655d4ba620b4257c2e07a6cb8c892cd293bb48aac81` | 任务专属重建脚本；硬编码冻结输入哈希、三图 ID/名称与逐图审计后的目标几何 |
| `adaptation/upstream-case.json` | 3358 | `d9f54a8a5ae4ea6b00ad35fb1425478a2b6fa557ba0f81ddf5f788aa0267258f` | Query、style target、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入和上游 Ground Truth 审计

冻结输入为 2 页、960×540pt 横版文稿，页面对象数为 `[2, 12]`，共 4 个媒体和 43 个 ZIP 包成员。第 2 页的三张图实际为：`Picture 26` 瀑布、`Picture 4` 螺旋楼梯、`Picture 18` 国际象；它们在输入中分别错位、越页、尺寸不一且两张有旋转，不能仅按对象名称猜测。

上游 Ground Truth 将三图大致放到正确位置，但同时改动 21 个包部件：替换 `theme1.xml`，重写 11 个 slide layout 和 1 个 slide master，并在封面增加 2 个渐变矩形和 1 条金色线。原尺寸渲染中，封面被大面积压暗，第 2 页 `Motion & Stills` 标题不可见，照片还遮挡了三个标签的大部分文字。上游 style target 也明确将这些换肤定义为工具副作用，因此它不符合本题保持契约，只作证据和负例。

## 本地 Golden 构造

本地 Golden 由制题者在完整读取三张图的内容、媒体哈希、裁剪、几何、旋转、标签位置和上游副作用后逐题重建，不是通用批量 mock：

- 从冻结输入出发，43 个包成员中只改 `ppt/slides/slide2.xml`。
- 在第 2 页显式写入主题 `bg1` 纯色背景，在原主题下渲染为白色。
- 瀑布、楼梯和国际象的左边位置分别为 75.00pt、354.39pt 和 632.83pt，顶边均为 135pt。
- 三图均设为 225×225pt，保留原媒体和原裁剪，旋转归零；照片底边与标签上边保留约 3.1–4.9pt 的可见间隙。
- 封面、主题、母版、版式、粉色面板、青蓝装饰、标签、标题、备注、4 个媒体和其他 42 个包部件字节保持。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，根据任务契约独立重算。

- 使用三张图真实媒体内容的 SHA256 识别语义身份，再检查瀑布—楼梯—国际象的从左到右顺序和目标水平位置，不依赖 Golden 文件。
- 逐图检查旋转、顶边、宽高、方形性、跨图尺寸差、水平间隙、粉色面板包含关系和与对应标签的垂直间隙。
- 接受主题 `bg1/lt1` 纯色或近白 `sRGB` 纯色，拒绝继承背景、渐变、图片或深色背景。
- 按 z-order、对象 ID、名称和规范化 OOXML 签名比对第 2 页全部 9 个非目标对象，拒绝改标题、改标签、移粉色面板或重染青蓝装饰。
- 独立比对封面对象、封面背景、备注、`themeElements`、slide master 和 11 个 slide layout，专门拒绝上游 Ground Truth 的无关换肤。
- 媒体哈希 Counter、包成员集、页数、对象数和三图裁剪独立检查，拒绝替换图片、压平、额外素材和额外页。
- LibreOffice 实际渲染冻结输入与候选文稿；封面整页像素相似度必须不低于 0.9995，遮罩三图旧新旋转边界框后第 2 页其余区域不低于 0.9985。
- 规范化掉合法图片几何和显式背景后的内容 XML 与关键包部件字节保持只占 3 分，不会将合法答案锁死为某一个 Golden 字节流。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地逐题重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立近白 `sRGB`、225.5pt、小幅平移答案 | 1.000000 | 通过 | 背景表达和三图几何不同于 Golden，但满足契约 |
| 独立 `bg1`、226pt、134pt 顶边答案 | 1.000000 | 通过 | 与 Golden 尺寸和顶边不同，标签仍完整可见 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.970000 | 中等 | 功能、对象和渲染正确，只有 3 分字节级保持失败 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺文件、无法打开或任务未完成 |
| 上游 Ground Truth | 0.000000 | 拒绝 | 无关换主题、改封面且照片遮挡标签 |
| 缺图 / 错顺序 / 交换图片内容 | 0.000000 | 拒绝 | 三个语义图像与标签匹配失败 |
| 顶边不齐 / 尺寸不一 / 残留旋转 / 遮挡标签 | 0.000000 | 拒绝 | 三联画几何或可读性失败 |
| 继承背景 / 渐变背景 / 黑色背景 | 0.000000 | 拒绝 | 显式近白纯色背景失败 |
| 改封面 / 改标签 / 移面板 / 改装饰 / 改标题 | 0.000000 | 拒绝 | 非目标内容或主题保持失败 |
| 改裁剪 / 额外对象 / 额外页 / 答案标记 | 0.000000 | 拒绝 | 可编辑性、对象库或交付卫生失败 |

包含 fresh replay 在内共 28 组正例、边界例和负例，全部与 `expectations.json` 一致。同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena016-oracle-photo-triptych-20260828`，trial 为 `pptarena-016-aligned-photo-tript__dtEpd52`，0 个异常；容器产物 SHA256 为 `5791250ffc46eaa7a95b388d86c246b09030ce476b430b84a170d7c496c97f05`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-016-aligned-photo-triptych
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-016-aligned-photo-triptych
    harbor run --path tasks/ppt/pptarena-016-aligned-photo-triptych --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 流程渲染并原尺寸查看冻结输入、上游 Ground Truth 和本地 Golden 的全部 2 页。
- 已原尺寸查看两个独立合法变体的第 2 页；三张图映射正确、等大对齐、裁剪正常，三个标签完整可见。
- `visual_qa.py` 对冻结输入、上游 Ground Truth、本地 Golden 和两个合法变体都报告 0 条启发式警告；上游的主题/遮挡问题是由逐页目视和专属 Judge 捕获，不依赖通用启发式。
- 本地 Golden 封面与冻结输入的渲染像素相似度为 1.0，遮罩三图旧新边界框后第 2 页其余区域为 0.999127。

## 剩余风险

- 正式公开 Release 前必须复核模板图与三张摄影图片的第三方权利；如无明确授权，需替换为可公开素材并保持相同的错位、旋转、裁剪和标签映射任务语义。
- LibreOffice/PDF 渲染可检查布局、标签和背景，但不能穷尽 Microsoft PowerPoint 的全部图片编辑行为；正式 Release 前仍建议在 PowerPoint 中手动拖动一张图并查看裁剪面板一次。
- 当前 L2 / Medium 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。
