# PPTArena Case 14 逐题适配记录

## 适配契约

- 角色：学术演示文稿编辑。
- 输入：34 页博弈论学术演示文稿，其中第 8、9、23、31 页为隐藏页。
- 真实意图：将第 2–34 页全部主标题设为粗体，同时保持第 1 页封面标题和整份文稿的其他内容。
- 主交付物：Bolded_BoldTitleTestA.pptx。
- 不可妥协约束：所有目标标题均为粗体，封面不变，标题的其他属性不变，正文、备注、媒体和隐藏状态不变，文件仍为原生可编辑 PPTX。
- 候选设计目标：L1 / Easy；正式难度字段保持 unassigned，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena Case 14: Bold Title (A)。
- GitHub：https://github.com/michaelofengenden/PPTArena.git。
- Git commit：6f26a63ec1014b98c44990c7b58133b832070be2。
- Hugging Face dataset commit：20fb8902c74df2e97a2c31d288e4fa7f23405157。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在 upstream-github-readme.md、upstream-hf-dataset-card.md 和 LICENSE-PPTArena-MIT.txt。
- 上游 Query、style target、输入和人工 Ground Truth 映射完整保存在 upstream-case.json；该 JSON 已与冻结索引逐字段比对。
- 上游数据卡明确称 Ground Truth 为人工编辑目标，本题不把输入或答案伪称为自行编造的数据。

原文稿为公开数据集中的学术演示材料。本候选题依据上游 MIT 数据集授权保留完整归因；在正式 Release 前仍需纳入项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件                                                        | 字节数 | SHA256                                                           | 实际检查                                                   |
|-------------------------------------------------------------|-------:|------------------------------------------------------------------|------------------------------------------------------------|
| data/input_files/BoldTitleTestA_Game_theory_4.pptx          | 1144147 | a8f00f2d2c814e5ef9c81a7757a72588ee911af67c7d120c5ca078d3a29395c1 | 34 页含隐藏页渲染、逐页原尺寸查看、标题对象和全部文字审计 |
| solution/Golden/Bolded_BoldTitleTestA.pptx                  | 1102313 | d2547d52b8237a46f41fad8c13a1e5b10322443def2f9c52d7611bee106c1947 | 34 页含隐藏页渲染、逐页原尺寸查看、OOXML 和媒体差异审计  |
| adaptation/upstream-case.json                               |     889 | dac2f7fe087c0685277f307eb740672a91086c3b94fbf10b6b1eb982203147cb | Query、style target 与冻结索引逐字段比对                  |
| adaptation/upstream-github-readme.md                        |    4576 | edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58 | 数据结构、规模、评测方式和许可标记                         |
| adaptation/upstream-hf-dataset-card.md                      |    5021 | 163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa | MIT 字段、公开状态、人工 Ground Truth 说明                 |
| adaptation/LICENSE-PPTArena-MIT.txt                         |    1068 | dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8 | 上游 MIT 许可文本                                          |

## 上游答案审计与适配决策

输入和 Golden 均为 34 页、10×7.5 英寸，共有 389 个文本元素，且每页都有可识别的主标题占位符。

- 第 1 页标题内容、样式和几何完全不变。
- 第 2–24、27、30–34 页共 29 页的标题被显式设置为粗体。
- 第 25、26、28、29 页的 Pluribus 标题原本已经是粗体，Golden 保持不变。
- Golden 中第 2–34 页每个非空标题文字运行最终都为粗体。
- 第 8、9、23、31 页的隐藏状态保持；它们的标题同样完成了粗体处理。
- 标题除粗体外的文字、字体、字号、颜色、斜体、下划线、段落格式和几何均保持。
- 全部非标题对象、原文、备注和 43 个媒体文件保持，未发现跨页移动或新增媒体。
- 第 19 页标题虽然在结构上新增了显式粗体属性，但其主题继承效果本来已呈现粗体，因此输入与 Golden 的栅格图相同；Judge 同时检查结构语义和视觉结果，避免漏掉这种只看像素无法识别的修改。

本地 Judge 不读取 Golden，也不比较 Golden 哈希。它从冻结输入独立定位每页主标题，检查第 2–34 页粗体、第 1 页保护、标题非粗体属性、标题占位符与几何、逐页原文、讲者备注、非标题对象、隐藏状态、图片和媒体集合，并把隐藏页临时展开后用 LibreOffice 渲染 34 页。视觉比较只遮罩标题可能发生字重变化的区域，标题外区域必须与输入高度一致。

## 实际验证结果

| 样例                                     | 实际 reward | 结果 | 关键原因                                                   |
|------------------------------------------|------------:|------|------------------------------------------------------------|
| 上游人工 Golden                          |    1.000000 | 通过 | 11 条判据全部通过                                          |
| fresh replay                             |    1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致              |
| 空交付                                   |    0.000000 | 拒绝 | 缺失指定交付物                                             |
| 损坏 PPTX                                |    0.000000 | 拒绝 | ZIP / OOXML 无法打开                                      |
| 未修改输入                               |    0.000000 | 拒绝 | 29 页目标标题未完成显式粗体                                |
| 漏改第 17 页标题                         |    0.000000 | 拒绝 | 目标标题不完整                                             |
| 漏改隐藏页 8、9、23、31                  |    0.000000 | 拒绝 | 隐藏页也属于处理范围                                       |
| 误将第 1 页封面设为粗体                  |    0.000000 | 拒绝 | 封面保护失败                                               |
| 将封面、正文和标题全部设为粗体           |    0.000000 | 拒绝 | 封面与非标题样式均被破坏                                   |
| 将第 12 页标题改为红色                   |    0.000000 | 拒绝 | 标题非粗体属性保护失败                                     |
| 改写第 20 页正文                         |    0.000000 | 拒绝 | 原文和非标题对象保护失败                                   |
| 取消第 8 页隐藏状态                      |    0.000000 | 拒绝 | 隐藏状态保护失败                                           |
| 第 10 页整页截图                         |    0.000000 | 拒绝 | 原生标题和正文缺失、媒体与非文本对象变化                   |
| 夹带 solution/Golden 标记                |    0.000000 | 拒绝 | 新增未授权文字与答案路径                                   |
| 标题正确但第 17 页下移 0.08 英寸         |    0.913580 | 中等 | 核心内容正确，仅标题几何判据失败                           |

同一 Golden 独立判分 5 次的 reward 为 [1, 1, 1, 1, 1]，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 最终回放结果为 1.0，job 为 pptarena014-oracle-boldtitles-v2-20260828，trial 为 pptarena-014-bold-titles-except__pPQqfuS，0 个异常；容器产物 SHA256 为 d2547d52b8237a46f41fad8c13a1e5b10322443def2f9c52d7611bee106c1947。

首轮 Harbor 回放为 0：Oracle 脚本在容器里错误回退到 /data/input_files，而 Harbor 实际将输入挂载到 /app/input_files，因而没有生成输出。修正容器默认输入路径后，第二轮回放达到 1.0；Judge、Rubric 和 Golden 均未为此放宽。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-014-bold-titles-except-cover
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-014-bold-titles-except-cover
    harbor run --path tasks/ppt/pptarena-014-bold-titles-except-cover --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已查看输入和 Golden 的 34 页联系表。
- 已把第 8、9、23、31 页临时展开，并逐页查看 Golden 的 34/34 页原尺寸渲染图。
- 已核对标题字重变化没有造成新增裁切、遮挡、空白页或离页内容。
- 官方 Harbor 容器内展开隐藏页后的 34 页均可渲染；标题外区域最低像素相似度为 1.0，封面相似度为 1.0。
- 原输入第 17 页标题框右侧本来就超出版心约 0.044 英寸，输入与 Golden 完全一致；Judge 按“不得新增越界”检查，避免把上游既有特征误判为本题新增错误。

## 剩余风险

- 正式公开 Release 前仍需对上游学术材料中的第三方图片和内容做项目级法务复核；当前依据是 PPTArena 公开数据卡和仓库的 MIT 标记。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过真实 Agent Rollout 后才能写回 difficulty_level 和 difficulty_tier。
