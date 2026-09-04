# PPTArena Case 15 逐题适配记录

## 适配契约

- 角色：课程演示文稿编辑。
- 输入：11 页人工智能课程演示文稿。
- 真实意图：把每一页的主标题设为粗体；第 1 页只处理主标题 CSE1300，不改课程副标题。
- 主交付物：Bolded_BoldTitleTestB.pptx。
- 不可妥协约束：11 页全部主标题均为粗体，标题的其他属性不变，课程副标题、正文、备注、校徽和版式不变，文件仍为原生可编辑 PPTX。
- 候选设计目标：L1 / Easy；正式难度字段保持 unassigned，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena Case 15: Bold Title (B)。
- GitHub：https://github.com/michaelofengenden/PPTArena.git。
- Git commit：6f26a63ec1014b98c44990c7b58133b832070be2。
- Hugging Face dataset commit：20fb8902c74df2e97a2c31d288e4fa7f23405157。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在 upstream-github-readme.md、upstream-hf-dataset-card.md 和 LICENSE-PPTArena-MIT.txt。
- 上游 Query、style target、输入和人工 Ground Truth 映射完整保存在 upstream-case.json；该 JSON 已与冻结索引逐字段比对。
- 上游 evaluator 使用初始文稿、Ground Truth 与候选文稿的结构化 smart diff，并用视觉模型比较渲染页；本次适配保留其语义目标，但改为可离线重放的任务专属确定性 Judge。
- 上游数据卡明确称 Ground Truth 为人工编辑目标，本题不把输入或答案伪称为自行编造的数据。

原文稿为公开数据集中的课程演示材料。本候选题依据上游 MIT 数据集授权保留完整归因；在正式 Release 前仍需纳入项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件                                                        | 字节数 | SHA256                                                           | 实际检查                                                    |
|-------------------------------------------------------------|-------:|------------------------------------------------------------------|-------------------------------------------------------------|
| data/input_files/BoldTitleTestB_cse1300-ai.pptx             | 166581 | 02687b50c86639c32501eefdf04cfb45e1687f45f9502c7896f4bf46f9e892b0 | 11 页渲染、逐页原尺寸查看、标题与封面文本对象审计          |
| solution/Golden/Bolded_BoldTitleTestB.pptx                  |  91524 | ff5b66801f4c7177a53b0fca2656d2660e363574b9233a35fa7bd43b11bb290a | 11 页渲染、逐页原尺寸查看、OOXML、备注和媒体差异审计       |
| adaptation/upstream-case.json                               |    862 | d5459b99d4bbe48b25332cc28bf682ae27fb86dd1acb678457e5870f0b3946bc | Query、style target 与冻结索引逐字段比对                   |
| adaptation/upstream-github-readme.md                        |   4576 | edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58 | 数据结构、规模、评测方式和许可标记                          |
| adaptation/upstream-hf-dataset-card.md                      |   5021 | 163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa | MIT 字段、公开状态、人工 Ground Truth 说明                  |
| adaptation/LICENSE-PPTArena-MIT.txt                         |   1068 | dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8 | 上游 MIT 许可文本                                           |

## 上游答案审计与适配决策

输入和 Golden 均为 11 页、10×7.5 英寸，共有 60 个文本段落。每页都有可识别的主标题占位符。

- 第 1 页主标题是 CSE1300；其下方 Introduction to Artificial Intelligence (AI) 是非目标课程副标题。
- 第 1–11 页所有主标题均在 Golden 中显式设为粗体；多运行和多行标题的每个非空文字运行都为粗体。
- 标题除粗体外的文字、字体、字号、颜色、斜体、下划线、段落格式和几何均保持。
- 全部非标题对象、逐页原文和讲者备注保持。
- 唯一一张校徽媒体的字节哈希和页面内位置保持；没有新增图片或整页压平图层。
- 输入与 Golden 的 ZIP 成员名一致。Golden 体积更小来自 Office 保存时的包压缩差异，不是内容或媒体丢失。

本地 Judge 不读取 Golden，也不比较 Golden 哈希。它从冻结输入独立定位每页主标题，检查 11 页粗体、封面 CSE1300 的语义定位、课程副标题保护、标题非粗体属性、占位符与几何、逐页原文、讲者备注、非标题对象、图片和媒体集合，并用 LibreOffice 渲染 11 页。视觉比较只遮罩标题可能发生字重变化的区域，标题外区域必须与输入高度一致。

## 实际验证结果

| 样例                                      | 实际 reward | 结果 | 关键原因                                                   |
|-------------------------------------------|------------:|------|------------------------------------------------------------|
| 上游人工 Golden                           |    1.000000 | 通过 | 11 条判据全部通过                                          |
| fresh replay                              |    1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致              |
| 空交付                                    |    0.000000 | 拒绝 | 缺失指定交付物                                             |
| 损坏 PPTX                                 |    0.000000 | 拒绝 | ZIP / OOXML 无法打开                                      |
| 未修改输入                                |    0.000000 | 拒绝 | 11 页主标题未完成粗体                                      |
| 漏改第 1 页 CSE1300                       |    0.000000 | 拒绝 | 封面主标题属于处理范围                                     |
| 漏改第 9 页标题                           |    0.000000 | 拒绝 | 目标标题不完整                                             |
| 第 6 页多运行标题只加粗部分文字           |    0.000000 | 拒绝 | 标题每个非空文字运行都必须为粗体                           |
| 取消第 1 页课程副标题粗体                 |    0.000000 | 拒绝 | 非主标题保护失败                                           |
| 将正文和课程副标题全部设为粗体            |    0.000000 | 拒绝 | 非标题样式保护失败                                         |
| 将第 8 页标题改为红色                     |    0.000000 | 拒绝 | 标题非粗体属性保护失败                                     |
| 改写第 10 页正文                          |    0.000000 | 拒绝 | 原文和非标题对象保护失败                                   |
| 第 4 页整页截图                           |    0.000000 | 拒绝 | 原生标题和正文缺失、媒体与非文本对象变化                   |
| 夹带 solution/Golden 标记                 |    0.000000 | 拒绝 | 新增未授权文字与答案路径                                   |
| 标题正确但第 10 页下移 0.08 英寸          |    0.914634 | 中等 | 核心内容正确，仅标题几何判据失败                           |

同一 Golden 独立判分 5 次的 reward 为 [1, 1, 1, 1, 1]，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 pptarena015-oracle-boldtitles-20260828，trial 为 pptarena-015-bold-all-slide-titl__QFHRUCk，0 个异常；容器产物 SHA256 为 ff5b66801f4c7177a53b0fca2656d2660e363574b9233a35fa7bd43b11bb290a。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-015-bold-all-slide-titles
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-015-bold-all-slide-titles
    harbor run --path tasks/ppt/pptarena-015-bold-all-slide-titles --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已逐页查看输入和 Golden 的 11/11 页原尺寸左右对照图。
- 已确认第 1 页只改变 CSE1300 的字重，课程副标题和校徽保持。
- 已确认第 2–11 页标题字重变化没有造成新增裁切、遮挡、空白页或离页内容。
- 官方 Harbor 容器内 11 页均可渲染，标题外区域最低像素相似度为 1.0。
- 原输入第 2、4、6、8、9 页已有正文接近或覆盖右下角校徽，视觉 QA 在输入与 Golden 上给出相同警告；本题要求保持，不把这些上游既有版式问题冒充本题新增缺陷。

## 剩余风险

- 正式公开 Release 前仍需对上游课程材料中的校徽和第三方内容做项目级法务复核；当前依据是 PPTArena 公开数据卡和仓库的 MIT 标记。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过真实 Agent Rollout 后才能写回 difficulty_level 和 difficulty_tier。

## 视觉评分适配与校准

制题模型已逐项读取 `instruction.md`、`task.toml`、`rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、上游 Case 记录、完整解题脚本、原始输入和人工 Golden。输入与 Golden 均已渲染 11 页联系图并逐页对照，另外以原始分辨率查看封面和多行标题第 6 页。本次检查确认 Golden 的第 1 页只将 `CSE1300` 处理为目标主标题，课程副标题、黑金双边框和右下校徽保持。

本题是精确语义定位与字重编辑任务，因此采用 `reference` 模式。专属视觉 Rubric 共 4 条：11 页主标题粗体的完整一致性、封面 `CSE1300` 与课程副标题的语义层级、多行长标题加粗后的换行/容纳，以及边框、校徽、正文和留白保真。代码 Judge 继续独立检查每个非空文字运行的粗体语义、副标题对象、非粗体属性、几何、备注、媒体和标题外像素保真；VLM 只补充实际可见的字重层级、换行、留白和专业完成度。原输入已存在的某些正文与右下校徽重叠仅用作参考基线，不因 Golden 保留而重复扣分。综合分按代码分 70% + VLM 分 30% 计算，两个分项单独保存。

Golden 的 5 次独立 Terra 视觉分数为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。

中等可见负例仅让第 8 页 `Common types of Artificial Neural Networks` 主标题丢失粗体层级，其余 10 页、黑金双边框、右下校徽、正文和留白仍完整，文件仍可打开且保持原生可编辑。其代码分为 0.0，Terra 视觉分为 0.860000，按 70/30 权重得到综合分 0.258000；视觉分严格低于 Golden，且高于空白或损坏文件。本题 6 次 Terra 调用均一次成功，未发生基础设施失败；若后续出现调用失败，仍必须标记为评测基础设施失败，不记为 0 分，也不退化为纯代码总分。
