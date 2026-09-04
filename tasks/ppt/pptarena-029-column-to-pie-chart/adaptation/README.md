# PPTArena Case 29 逐题适配记录

## 适配契约

- 角色：健康政策数据演示文稿编辑。
- 输入：9 页 Kaiser Family Foundation 亚裔与 NHOPI 健康政策演示文稿。
- 真实意图：将第 3 页现有簇状柱形图转换为原生可编辑饼图。
- 主交付物：`PieChart_GroundTruthA.pptx`。
- 不可妥协约束：保留单系列 `%`、8 个族群类别、类别—数值映射和底层小数；全部 8 类可辨识且数值或百分比可读；除目标图表外不改动原文稿。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 29: Convert Bar Chart to Pie Chart。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入、人工 Ground Truth 路径和宽松布局约束已保存在 `upstream-case.json`，并与冻结索引逐字段比对。
- 上游 evaluator 通过 smart JSON diff 和视觉模型判断候选文稿，明确 Ground Truth 只是一个有效实现，不强制精确坐标。本次完整阅读了 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，但未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开演示文稿和人工 Ground Truth，不把输入、Query 或答案伪称为本项目自行编造的数据。

原文稿包含公开机构标识、健康统计数据和第三方媒体。本候选题依据 PPTArena 的公开数据集和 MIT 标记保留归因；正式 Release 前仍需做项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/PieChart_TestA.pptx` | 322109 | `07911f768a13823f9e8a4a426a7d2667a16f8283a4237f2883abe8b24527415e` | 9 页逐页渲染、原尺寸观察、图表数据、对象、备注和媒体审计 |
| `solution/Golden/PieChart_GroundTruthA.pptx` | 323524 | `1bc36f7e0165538765a77b9d8ddd5df473afdca0babe95cda229daa1c286c36c` | 9 页逐页渲染、原尺寸观察、OOXML 和输入差异审计 |
| `adaptation/upstream-case.json` | 869 | `423212717aac0655a83bc6a30b0f3269bd5189226534174ea3ce9c8741fb82b6` | Query、style target、路径与改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| 上游 `src/evaluation_pairs_refined.json` | 165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 29 原始记录和不锁定精确坐标的等价答案要求 |
| 上游 `src/llm/judge.py` | 40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分调用链 |
| 上游 `agent_bench/judge_predictions.py` | 16100 | `3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样判分和结果汇总入口 |

## 文稿与 Golden 逐页审计

输入和 Golden 都是 9 页、10×7.5 英寸，共 35 个形状、25 个文字对象、8 张原生图表、1 个页面图片对象和 11 个媒体文件，无宏。

- 第 3 页原图表是 `COLUMN_CLUSTERED`，人工 Golden 将其改为 `PIE`；两者都是原生 Chart 对象。
- 图表坐标为 x=0.1007、y=1.9167、w=9.7986、h=4.8700 英寸；人工 Golden 保持该主内容区几何。
- 单系列名为 `%`；8 个类别为 Asian Indian、Chinese、Filipino、Japanese、Korean、Vietnamese、Other Asian 和 NHOPI。
- 底层数值为 `[0.228282261236966, 0.215493168539373, 0.149576327518770, 0.0337280980154271, 0.0767151130048191, 0.0990852974196732, 0.165631750738918, 0.031487983526053]`，人工 Golden 逐值保持。
- 人工 Golden 为 8 个类别显示类别名和数值标签，明确标签字号为 16pt，无图例。
- 其他 8 页 LibreOffice 渲染逐页完全一致，第 3 页图表区外渲染相似度为 1.0；11 个原媒体哈希全部保留。
- 已原尺寸目视检查输入/Golden 的 9/9 组左右对照图，只有第 3 页目标图表发生任务要求内的视觉变化。
- `visual_qa.py` 对输入和 Golden 都报告相同的 4 条 `underfilled_textbox` 通用警告，分别位于第 1、3、7、8 页；本题不擅自修改上游既有版式。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，并根据任务契约独立重算。

- 候选文稿必须严格为 9 页且页面尺寸不变；ZIP 包可打开、无损坏成员和宏。
- 第 3 页必须恰有一张 `PIE` 或 `PIE_EXPLODED` 原生图表，不允许柱形图、折线图、第二张图表或整页截图。
- 类别和数值直接从冻结输入的 Chart OOXML 缓存和 Chart 对象中解析；允许调整类别顺序，但类别—数值映射、系列名和底层数值必须一致。
- 数据标签同时支持 PowerPoint 写入的 series-level `dLbls` 和 python-pptx 写入的 plot-level `dLbls`，防止把结构不同但语义正确的饼图误拒。
- 全部 8 个类别必须通过图例或类别标签可辨识，且全部显示数值或百分比；显式标签字号不得小于 8pt。该项是核心门槛。
- 饼图必须在主内容区，宽度不小于 4 英寸、高度不小于 3 英寸；布局不锁定 Golden 坐标。
- 非目标页面通过工具中立的内容、样式、几何、图片哈希、图表数据和备注签名保护；第 3 页的全部非图表对象单独保护。
- LibreOffice 实际渲染冻结输入和候选文稿，非目标页最低像素相似度阈值为 0.9985；第 3 页忽略输入和候选图表包围盒的并集后，其他区域阈值也是 0.9985。
- 独立合法变体从冻结输入出发，新建 8.4×4.75 英寸饼图，使用右侧图例和百分比标签，不复制人工 Golden 的坐标、无图例选择或类别+数值标签形式。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 上游人工 Golden | 1.000000 | 通过 | 12 条判据全部通过 |
| 独立布局和标签的合法替代答案 | 1.000000 | 通过 | 右侧图例+百分比标签，不依赖单一 Golden 布局 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或仍是柱形图 |
| 错误图表类型 / 第二张图表 | 0.000000 | 拒绝 | 未满足唯一原生饼图门槛 |
| 篡改数值 / 篡改类别 | 0.000000 | 拒绝 | 冻结输入中的类别—数值映射校验失败 |
| 改写第 3 页主标题 / 改写第 5 页图表 | 0.000000 | 拒绝 | 目标页非图表内容或非目标页保持失败 |
| 用整页截图压平第 3 页 | 0.000000 | 拒绝 | 缺失原生饼图且整页图片覆盖 |
| 夹带 `solution/Golden` 标记 | 0.000000 | 拒绝 | 未授权文字、非图表内容和渲染保持同时失败 |
| 删除图例与标签 / 将标签改为 4pt | 0.000000 | 拒绝 | 8 类可辨识和可读数值核心门槛失败 |
| 将饼图缩小为 3.5×2.5 英寸 | 0.920000 | 中等 | 内容、数据和可编辑性正确，仅尺寸分项失败 |
| 错误新增第 10 页 | 0.000000 | 拒绝 | 页数和原文稿结构变化 |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena029-oracle-column-to-pie-20260828`，trial 为 `pptarena-029-column-to-pie-chart__RcYQpHq`，0 个异常；容器产物 SHA256 为 `1bc36f7e0165538765a77b9d8ddd5df473afdca0babe95cda229daa1c286c36c`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-029-column-to-pie-chart
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-029-column-to-pie-chart
    harbor run --path tasks/ppt/pptarena-029-column-to-pie-chart --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 的流程使用 `render_slides.py` 渲染输入、人工 Golden 和独立合法变体。
- 已逐页查看输入/Golden 的 9/9 组原尺寸对照图，并原尺寸查看独立合法变体第 3 页。
- 独立合法变体的 8 个饼图切片、百分比标签和右侧图例均可见，未与标题、小标题、注释、来源或页脚重叠。
- 人工 Golden 与 fresh replay 均可由 LibreOffice 渲染为完整 9 页；饼图区灰度标准差为 62.4347，非近白像素比例为 0.2285，不是空图或裁切图。

## 剩余风险

- 正式公开 Release 前仍需对原文稿中的 Kaiser Family Foundation 标识、健康统计数据和 11 个媒体文件做项目级法务与归因复核。
- 当前已通过 OOXML、python-pptx 语义对象、LibreOffice 渲染和人工对照；正式 Release 前仍建议用 Microsoft PowerPoint 再做一次打开/保存抽检。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case 与任务专属 `solve.py`，并核对第 3 页原柱形图的八类数据、固定文本和其余八页保持约束。
- 冻结输入与上游人工 Golden 均已用 LibreOffice 完整渲染、制作九页联系图并按原始分辨率查看，尤其放大比较第 3 页。Golden 用一张占据主内容区的蓝灰饼图替代原柱形图，八类和百分比均可见，其他页面没有视觉变化。
- 上游任务明确允许不同饼图位置、图例或标签方案，只要数据和视觉表达正确，因此使用 `reference_free`。`V001` 查饼图主构图，`V002` 查八类及百分比可读性，`V003` 查 KFF 蓝灰配色和切片区分，`V004` 查第 3 页固定标题/来源/页脚保持，`V005` 查其余八页与交付卫生。原生 Chart 类型、底层数据、标签开关、媒体与可编辑性继续由确定性 Judge 负责。
- Golden 的可见风险是 Japanese、Korean、Vietnamese 与 NHOPI 等较小切片的外部标签相对密集；VLM 应按实际渲染判断是否仍能无歧义辨识，不能因为底层数据正确而忽略标签碰撞。
- Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。中等可见负例保留了唯一原生饼图、八类底层数据、蓝灰配色、第 3 页标题/注释/来源/KFF 标识和其余八页，但将饼图缩小并下移，使右侧 `Chinese` 类别及其百分比在实际渲染中不再可见；文件仍可打开、数据可编辑，确定性代码分仍为 1.0。Terra 视觉分为 0.837500，按 70/30 权重得到综合分 0.951250；视觉分严格低于 Golden，且高于空白或损坏文件。本题 6 次 Terra 调用均一次成功，未发生基础设施失败；后续评分调用失败仍只记为基础设施失败，不按 0 分计入。
