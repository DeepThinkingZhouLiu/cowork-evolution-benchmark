# PPTArena Case 27 逐题适配记录

## 适配契约

- 角色：图文演示文稿编辑。
- 输入：1 页 16:9 PPTX，共 7 个原生对象，其中 3 个是已填充鲨鱼、汽车和棒球照片的图片占位符。
- 真实意图：根据图片实际内容将三图与 `shark/car/baseball` 标签正确配对，并整理为等大、水平对齐、间距均衡的三列布局。
- 主交付物：`FixImagePorportions_GroundTruthA.pptx`。
- 不可妥协约束：图片语义对应、原媒体和裁剪保持；标题、标签、背景、主题、母版、12 个版式、备注和超链接保持；三图保持原生可编辑。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 27: Correct Images。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>；commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本已保存。
- 上游 Query、style target、输入盘点、三张图的媒体哈希和适配边界逐字段保存在 `upstream-case.json`。

三张摄影图片来自外部网页或模板，其中文稿关系还保留了 PxHere 链接。PPTArena 数据集的 MIT 标记不会自动消除第三方图片权利风险，正式 Release 前必须单独复核或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/FixImagePorportions_TestA.pptx` | 4373662 | `026285dad008ed551128b04548909d8e4377324b67f8428cb6a68d2c87d9555f` | 1 页原尺寸渲染；7 个对象、3 个媒体、42 个包成员、逐图语义、裁剪、几何、标签、z-order 和超链接审计 |
| `solution/Golden/FixImagePorportions_GroundTruthA.pptx` | 4371193 | `0d7ae976a9601c6a803d7b64220ea972fcb059be85dc4b263dafa04b97727206` | 从冻结输入重建；原尺寸目视、OOXML 差异、媒体/裁剪哈希和 LibreOffice 打开渲染审计 |
| `adaptation/upstream-ground-truth-layout-mutated.pptx` | 4372761 | `9a9da8c820d825e75a144c0232563b5e5dbbf6cc63984dda6a1547a09897e469` | 上游人工 Ground Truth；目视布局正确，但确认无关重写母版和 11 个版式 |
| `adaptation/build_golden.py` | 3008 | `de8c83246290ba4ba3f85ad6aed72266c2ecd5552c044e314b02b423f9fef450` | 任务专属重建脚本；冻结输入哈希、3 个图片对象 ID 与逐图目标几何显式写入 |
| `adaptation/upstream-case.json` | 3670 | `01afa6b9b12280a68395df91db0e89937ee635cca8799477d7edf0e27627649c` | Query、style target、来源版本、媒体语义对应与不沿用的上游副作用逐项比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入与上游 Ground Truth 审计

冻结输入是 960×540pt 的单页文稿，标题为 `What I like`，三个标签位于 y=413.96pt，宽度均为 229.88pt。逐图目视与媒体哈希审计确定：shape 10 / image1.png 为鲨鱼，shape 4 / image2.jpeg 为汽车，shape 7 / image3.jpg 为棒球。原页面将棒球放在左上、鲨鱼放在中间且过大、汽车放在右上，图片覆盖标题且与标签全部错配。

上游 Ground Truth 的目视布局正确：鲨鱼、汽车、棒球从左到右，三图均为 229.88×229.88pt，顶边 y=170.34pt，与标签约有 13.74pt 空隙。但它共改动 18 个包部件：除文档属性、缩略图、演示文稿和视图状态外，还重写 `slideMaster1.xml` 与 `slideLayout1–11.xml`。它通过删除汽车/棒球图的显式 `xfrm`、再改全局版式占位符来实现对齐，违反局部编辑和最小改动契约，因此只作设计证据和负例。

## 本地 Golden 构造

本地 Golden 不是通用批量 mock，而是制题者在逐图识别媒体内容、标签中心、图片裁剪、上游视觉方案和全局副作用后从冻结输入重建：

- 42 个包成员中只改 `ppt/slides/slide1.xml`。
- 鲨鱼图 x=86.40pt，汽车图 x=366.87pt，棒球图 x=647.33pt；三图 y=170.34pt，尺寸均为 229.88×229.88pt，旋转为 0。
- 图片与标签中心偏差分别约为 0.00pt、0.59pt 和 1.19pt，纵向空隙均为 13.74pt；两个水平间距为 50.59pt 和 50.58pt。
- 3 个原媒体、裁剪、超链接、标题、标签、主题、母版、12 个版式、备注和其他 41 个包部件字节保持。

## Judge 设计

本地 Judge 不读取 Golden，也不比较 Golden 哈希；只从候选 PPTX 和冻结输入独立重算。

- 使用三张图真实媒体内容的 SHA256 识别鲨鱼、汽车和棒球，检查左到右语义顺序，不依赖对象名或 Golden。
- 检查三图宽高范围、方形性、跨图尺寸差、顶边差、旋转、页面边界、两个水平间距及其差值。
- 独立获取 `shark/car/baseball` 标签几何，检查图标中心偏差、图标纵向空隙与跨列一致性。
- 对比 4 个非图片对象的 z-order、ID、名称和规范化 OOXML 签名，拒绝改标题、改标签或移标签。
- 比对媒体哈希 Counter、图片裁剪、包成员集、关系、主题、母版、12 个版式和备注，专门拒绝上游 Ground Truth 的全局版式重写。
- LibreOffice 实际渲染冻结输入与候选文稿；遮罩三图旧新矩形后，其他区域像素相似度必须不低于 0.9980。
- 规范化掉三图合法几何后的内容与包骨架保持只占 2 分，合法答案可使用不同于 Golden 的尺寸和平移。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地逐图重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 224pt、y=174pt 的独立居中答案 | 1.000000 | 通过 | 几何不同于 Golden，但语义、等大、对齐与空隙全部合法 |
| 232pt、y=165pt 的独立居中答案 | 1.000000 | 通过 | 使用另一尺寸和顶边，仍不遮标题/标签 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.980000 | 中等 | 功能、媒体、布局和渲染正确，只有 2 分包骨架字节保持失败 |
| 空交付 / 损坏 PPTX / 未修改原件 | 0.000000 | 拒绝 | 缺失、无法打开或语义/几何任务未完成 |
| 上游 Ground Truth | 0.000000 | 拒绝 | 布局正确，但无关改母版和 11 个版式 |
| 缺图 / 图序错 / 交换媒体内容 | 0.000000 | 拒绝 | 三个语义媒体与标签匹配失败 |
| 尺寸不一 / 非方形 / 顶边不齐 / 旋转 | 0.000000 | 拒绝 | 三列几何契约失败 |
| 遮挡标签 / 图标偏心 / 间距不均 | 0.000000 | 拒绝 | 图标空间关系或可读性失败 |
| 改裁剪 / 改标签 / 移标签 / 改标题 | 0.000000 | 拒绝 | 图片内容或非目标对象保持失败 |
| 改媒体 / 改主题 / 改版式 | 0.000000 | 拒绝 | 文稿资产和全局结构保持失败 |
| 压平三图 / 额外对象 / 额外页 / 答案标记 | 0.000000 | 拒绝 | 可编辑性、对象集、页数或交付卫生失败 |

`expectations.json` 共记录 30 组 Golden、独立正例、fresh replay、中等样例和负例，实际结果全部与期望一致。同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena027-oracle-correct-image-placeholders-20260828`，trial 为 `pptarena-027-correct-image-place__C7DNEZN`，0 个异常；容器产物 SHA256 为 `0d7ae976a9601c6a803d7b64220ea972fcb059be85dc4b263dafa04b97727206`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-027-correct-image-placeholders
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-027-correct-image-placeholders
    harbor run --path tasks/ppt/pptarena-027-correct-image-placeholders --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --jobs-dir /tmp/cowork-harbor-jobs --job-name pptarena027-oracle-correct-image-placeholders-20260828

## 视觉复核

- 已按 presentation-skill 流程渲染并原尺寸查看冻结输入、上游 Ground Truth 和本地 Golden 的全部 1 页。
- 已原尺寸查看 224pt 与 232pt 两个独立合法变体；三图语义正确、无拉伸，标题和三个标签完整可见。
- `visual_qa.py` 对冻结输入、上游 Ground Truth、本地 Golden 和两个合法变体均只报告原文稿已存在的标题文本框过高启发式警告；该标题和文本框未被修改。
- 本地 Golden 遮罩三图旧新矩形后的渲染相似度为 0.999600，高于 0.9980 保持阈值。

## 剩余风险

- 正式公开 Release 前必须复核 3 张摄影图片和页面模板的第三方权利；如无明确授权，需替换为可公开素材并保持相同的语义错配与几何错位。
- LibreOffice/PDF 渲染可检查图片内容、裁剪、布局和标签可读性，但不能穷尽 PowerPoint 对图片占位符和模板继承的全部行为；正式 Release 前建议在 PowerPoint 中手动拖动一张图并查看裁剪面板一次。
- 当前 L2 / Medium 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、任务专属 `build_golden.py` / `solve.py`，并核对三个输入媒体的语义与几何审计。
- 冻结输入、本地 Golden 与上游布局变体均已用 LibreOffice 完整渲染并按原始分辨率查看。输入中棒球遮住标题、鲨鱼居中且过大、汽车位于右侧；Golden 将鲨鱼、汽车、棒球整理为等大的左中右三列，而上游变体虽视觉正确却无关改写母版和版式。
- 由于指令允许约 230pt 的合理尺寸和位置浮动，不要求复刻单一 Golden 坐标，本题使用 `reference_free`。`V001` 查图文语义配对，`V002` 查等大方形与顶边，`V003` 查横向等距节奏，`V004` 查图片—标签中心及纵向留白，`V005` 查标题、标签、背景与局部编辑完成度。媒体哈希、裁剪、原生可编辑性、主题/母版/备注保持继续由确定性 Judge 负责。
- Golden 的可见风险是三张摄影图片本身的亮度与主体占比不同；VLM 应判断布局是否用同尺寸和自然裁剪建立均衡，而不能把照片内容差异误当作几何不一致。
- Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、the configured VLM provider 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。首轮负例对中间汽车图的缩小/位移过轻，the VLM evaluator 仍给出 1.0，因此拒绝该负例。合格 v2 负例保留鲨鱼—SHARK、汽车—CAR、棒球—BASEBALL 的正确映射、方形裁切、标题、标签和蓝灰主题，但将中间汽车图缩到两侧的约 70% 并向右下偏移，使顶边、尺寸、列间距和图标留白明显不一致；文件仍可打开且保持原生可编辑。其代码分为 0.0，the VLM evaluator 视觉分为 0.695000，按 70/30 权重得到综合分 0.208500；视觉分严格低于 Golden，且高于空白或损坏文件。本题 7 次 the VLM evaluator 调用均一次成功，未发生基础设施失败；后续评分调用失败仍只记为基础设施失败，不按 0 分计入。
