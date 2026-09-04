# PPTArena Case 21 逐题适配记录

## 适配契约

- 角色：学术演示文稿编辑。
- 输入：15 页 DSI 学生组织公开演示文稿。
- 真实意图：只把每页语义主标题改成 Title Case，同时保留缩写、正文、备注、媒体和全部版式。
- 主交付物：FixTitleCase_GroundTruthA.pptx。
- 不可妥协约束：15 页标题均正确；DSI 保持全大写；第 2、4 页接受上游明确允许的两种问句写法；所有非标题内容和原生可编辑结构保持。
- 候选设计目标：L1 / Easy；正式难度字段保持 unassigned，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena Case 21: Title Case Titles。
- GitHub：https://github.com/michaelofengenden/PPTArena.git。
- Git commit：6f26a63ec1014b98c44990c7b58133b832070be2，工作树干净。
- Hugging Face dataset commit：20fb8902c74df2e97a2c31d288e4fa7f23405157。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在 upstream-github-readme.md、upstream-hf-dataset-card.md 和 LICENSE-PPTArena-MIT.txt。
- 上游 Query、style target、输入和人工 Ground Truth 映射完整保存在 upstream-case.json；该 JSON 已与冻结索引逐字段比对。
- 上游 evaluator 先比较原件、Ground Truth 与候选文件的结构化 smart diff，再用视觉模型判断渲染结果；本次实际读过 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，但没有执行上游网络模型 Judge。
- 上游数据卡说明 Ground Truth 是人工制作，本题不把输入、Query 或答案伪称为本项目自行编造的数据。

原文稿公开显示学生组织、作者姓名、邮箱和社交账号。本候选题依据 PPTArena 的公开数据集与 MIT 标记保留归因；正式 Release 前仍需再次完成隐私和第三方图片法务复核。

## 已读入上下文的文件

| 文件                                               |   字节数 | SHA256                                                           | 实际检查                                                   |
|----------------------------------------------------|---------:|------------------------------------------------------------------|------------------------------------------------------------|
| data/input_files/FixTitleCase_TestA.pptx           |  1052351 | 5462ea413b6ecdf5ee13b0893c360e744886283949fed7a8b44fd78ed01e727f | 15 页渲染、逐页原尺寸查看、标题和全部文字对象审计           |
| solution/Golden/FixTitleCase_GroundTruthA.pptx     |  1071295 | 4b596d4f4fd84db6b331e7d06161643d7726e1c66979da393ecef780b65e1d89 | 15 页渲染、逐页原尺寸查看、OOXML、备注、媒体和标题差异审计  |
| adaptation/upstream-case.json                      |     1455 | 2d1c8bbdcc3d5a806d67b3616f63ae7d6069b2c0c28d2a8531ebcc702572a786 | Query、style target 与冻结索引逐字段比对                    |
| adaptation/upstream-github-readme.md               |     4576 | edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58 | 数据结构、规模、评测方式和许可标记                           |
| adaptation/upstream-hf-dataset-card.md             |     5021 | 163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa | MIT 字段、公开状态、人工 Ground Truth 说明                   |
| adaptation/LICENSE-PPTArena-MIT.txt                |     1068 | dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8 | 上游 MIT 许可文本                                            |
| 上游 src/evaluation_pairs_refined.json             |   165586 | 17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f | Case 21 原始记录与两种问句答案规则                           |
| 上游 src/llm/judge.py                              |    40829 | 82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2 | smart diff、Ground Truth 和视觉评分调用链                    |
| 上游 agent_bench/judge_predictions.py              |    16100 | 3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308 | Agent 产物定位、逐题评分和结果汇总入口                        |

## 逐页标题审计

前 14 页只有普通文本框，第 15 页才有真正的标题占位符。因此，本题 Judge 不能只检查 `slide.shapes.title`。下面的目标由输入与人工 Ground Truth 的对象级差异、逐页渲染和人工阅读共同确认。

| 页 | 输入标题                                                               | 允许的结果                                                                 |
|---:|------------------------------------------------------------------------|----------------------------------------------------------------------------|
|  1 | data science and informatics: creating student led data science curricula | Data Science and Informatics: Creating Student Led Data Science Curricula |
|  2 | who are we?                                                            | Who are we? 或 Who Are We?                                                 |
|  3 | brief history of dsi                                                   | Brief History of DSI                                                       |
|  4 | what do we do?                                                         | What do we do? 或 What Do We Do?                                           |
|  5 | our workshops:                                                         | Our Workshops:                                                             |
|  6 | workshop Coordination                                                  | Workshop Coordination                                                      |
|  7 | competing data science truths                                          | Competing Data Science Truths                                              |
|  8 | Workshop design                                                        | Workshop Design                                                            |
|  9 | workshop instructors                                                   | Workshop Instructors                                                       |
| 10 | workshop instructors                                                   | Workshop Instructors                                                       |
| 11 | results of workshops                                                   | Results of Workshops                                                       |
| 12 | growth of dsi                                                          | Growth of DSI                                                              |
| 13 | Growth Statistics                                                      | Growth Statistics                                                          |
| 14 | results of growth                                                      | Results of Growth                                                          |
| 15 | contact                                                                | Contact                                                                    |

输入和 Golden 均为 15 页、10×7.5 英寸，共 32 个文本对象、119 个文本段落、4 个页面图片对象和 5 个媒体文件，没有隐藏页和宏。第 1、6、7、15 页存在非空讲者备注，均已逐页比对并保留。

- Golden 只在上述 15 个标题中改变字母大小写；第 13 页原本已正确，没有改动。
- 第 2、4 页人工 Ground Truth 使用普通问句大小写；上游 refined style target 明确要求同时接受完整标题式写法。
- 第 3、12 页的 DSI 保持全大写。
- 标题字体、字号、颜色、粗体、斜体、下划线、段落设置、文本框类型和几何保持。
- 全部非标题文字、备注、页面顺序、媒体哈希和非文本对象保持。
- Office 保存人工 Ground Truth 时合并了少量等价文字运行，并重写嵌入字体子集和文档元数据；对象级样式、逐页内容和渲染结果没有因此改变。

## Judge 设计

本地 Judge 不读取 Golden，也不比较 Golden 文件哈希。它只读取冻结输入，并按逐页审计得到的标题文本与几何定位 15 个语义标题，不依赖形状 ID 或对象名称。

- 标题内容检查使用逐页允许集合；第 2、4 页各有两个合法答案。
- 标题样式检查忽略大小写和等价文字运行合并，但检查字体、字号、颜色、字重、斜体、下划线、语言、段落设置和文本框属性。
- 非标题文字、备注、非标题对象、图片、媒体哈希、页数、页面尺寸和原生可编辑结构均从冻结输入独立重算。
- LibreOffice 渲染全部 15 页。视觉比较只遮罩标题字形可能变化的区域；第 1 页原始标题已经溢出文本框并与作者行重叠，所以遮罩包含该既有溢出区域，正文仍由对象级判据严格保护。
- 新增整页图片、删除标题、夹带答案路径、额外页面和非标题改写都会被拒绝。

## 实际验证结果

| 样例                                      | 实际 reward | 结果 | 关键原因                                                     |
|-------------------------------------------|------------:|------|--------------------------------------------------------------|
| 上游人工 Golden                           |    1.000000 | 通过 | 11 条判据全部通过                                            |
| 合法替代答案 Who Are We? / What Do We Do? |    1.000000 | 通过 | 上游规则明确允许，不依赖单一 Ground Truth 字符串             |
| fresh replay                              |    1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致                |
| 空交付                                    |    0.000000 | 拒绝 | 缺失指定交付物                                               |
| 损坏 PPTX                                 |    0.000000 | 拒绝 | ZIP / OOXML 无法打开                                         |
| 未修改输入                                |    0.000000 | 拒绝 | 多页标题未完成 Title Case                                    |
| 漏改第 11 页                              |    0.000000 | 拒绝 | 仍为 results of workshops                                    |
| 将第 12 页写成 Growth of dsi              |    0.000000 | 拒绝 | DSI 缩写未保留                                               |
| 将第 2 页写成 Who Are we?                 |    0.000000 | 拒绝 | 不属于两种允许的问句写法                                     |
| 将第 14 页写成 Results Of Growth          |    0.000000 | 拒绝 | 短介词 of 的大小写错误                                       |
| 将第 8 页标题改为红色                     |    0.000000 | 拒绝 | 标题非大小写样式保护失败                                     |
| 改写第 2 页正文                           |    0.000000 | 拒绝 | 非标题原文和对象保护失败                                     |
| 删除第 10 页标题                          |    0.000000 | 拒绝 | 语义标题和原生文本对象缺失                                   |
| 第 10 页整页截图                          |    0.000000 | 拒绝 | 原生标题和正文缺失、媒体与非文本对象变化                     |
| 夹带 solution/Golden 标记                 |    0.000000 | 拒绝 | 新增未授权文字与答案路径                                     |
| 错误增加第 16 页                          |    0.000000 | 拒绝 | 页数和页面结构变化                                           |
| 标题正确但第 10 页下移 0.08 英寸          |    0.914634 | 中等 | 核心内容正确，仅标题几何判据失败                             |

同一 Golden 独立判分 5 次的 reward 为 [1, 1, 1, 1, 1]，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena021-oracle-titlecase-20260828`，trial 为 `pptarena-021-title-case-titles__bqjR2DP`，0 个异常；容器产物 SHA256 为 `4b596d4f4fd84db6b331e7d06161643d7726e1c66979da393ecef780b65e1d89`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-021-title-case-titles
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-021-title-case-titles
    harbor run --path tasks/ppt/pptarena-021-title-case-titles --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已用 presentation-skill 的 render_slides.py 渲染输入和 Golden，各 15 页。
- 已逐页查看输入和 Golden 的 15/15 页原尺寸左右对照图。
- 已确认所有标题的大小写变化，没有新增裁切、空白页、离页内容或正文变化。
- visual_qa.py 对输入和 Golden 均未报告新增警告。
- 官方 Harbor 容器内 15 页均可渲染；完成结构保护后，标题外最低像素相似度为 1.0。
- 第 1 页标题与作者行重叠是冻结输入和上游 Golden 中共同存在的原始版式问题，本题只做大小写编辑，不擅自重排。

## 剩余风险

- 正式公开 Release 前仍需对公开联系方式、学生组织材料和 5 个第三方媒体文件做项目级隐私与法务复核。
- 上游 Ground Truth 的保存过程重写了嵌入字体子集；当前已通过对象级字体检查、逐页人工对照和 LibreOffice 回放，但仍建议在正式 Release 前用 Microsoft PowerPoint 再做一次打开/保存抽检。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 difficulty_level 和 difficulty_tier。

## 任务专属视觉评分补充

制题模型重新完整读取了 instruction、task.toml、deterministic rubric、judge、expectations、solution、upstream case 与全部适配记录，并将冻结输入和 Golden 的 15 页全部渲染为联系图；另外放大检查了封面长标题、第 6、7 页多行标题、第 12 页 `DSI` 标题和第 15 页居中联系页。视觉评分采用 `reference` 模式，因为本题是对冻结演示稿进行边界明确的局部文字编辑，除允许的 Title Case 变化外，应直接以人工参考交付约束标题样式、正文、图像和整套版式的保持。

`visual_rubric.json` 为本题单独定义四个维度：15 页标题的 Title Case 与 `DSI`/问句特例、标题原字号/颜色/换行/位置及无新增溢出、蓝白渐变/橙色底带/页码/标识/正文/照片的逐页保持，以及无截图压平、遮罩或调试痕迹。评分锚点均给出 0、0.5、1 的可见证据；封面标题与作者行在冻结输入和 Golden 中共同存在的既有重叠只作为参考基线，候选未进一步恶化时不重复扣分。

Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。

首轮负例只让少量页面的标题改为棕红色并轻微漂移，Terra 仅识别出第 7 页颜色差异，视觉分为 0.9375，因区分力偏弱而拒绝。合格 v2 负例保留 15 页 Title Case 文本、`DSI` 缩写、背景、橙色底带、页码、照片、正文和联系方式，但在第 5、7、11、14、15 页制造标题棕红色、字号/位置偏移，并让第 11 页标题与首条正文产生明显重叠；文件仍可打开且保持原生可编辑。其代码分为 0.0，Terra 视觉分为 0.734000，按 70/30 权重得到综合分 0.220200；视觉分严格低于 Golden，且高于空白或损坏文件。本题未发生基础设施失败；任何后续 API、渲染或解析失败仍只记为评测基础设施失败，不计作 0 分。
