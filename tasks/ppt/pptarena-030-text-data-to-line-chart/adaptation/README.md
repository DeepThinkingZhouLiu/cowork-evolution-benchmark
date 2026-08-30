# PPTArena Case 30 逐题适配记录

## 适配契约

- 角色：健康政策数据演示文稿编辑。
- 输入：9 页 Kaiser Family Foundation 亚裔与 NHOPI 健康政策演示文稿。
- 真实意图：将第 8 页 2013–2017 年三个人群的明文未参保率转为原生可编辑、带数据点标记的折线图，并制作右侧自定义图例。
- 主交付物：`LineGraph_GroundTruthA.pptx`。
- 不可妥协约束：五个年份、三个系列和 15 个输入可见数值一致；折线、标记、坐标轴与自定义图例可编辑；隐藏默认图例；除目标区域外不改动原文稿。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 30: Convert Text Data to a Line Chart。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、完整 style target、输入、人工 Ground Truth 路径和输出边界已保存在 `upstream-case.json`，并与冻结索引逐字段比对。
- 上游 evaluator 使用 smart JSON diff 和视觉模型判断候选文稿，Ground Truth 只是一个参考实现。本次完整阅读了 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，但未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开演示文稿和人工任务意图，不把输入、Query 或参考答案伪称为本项目自行编造的数据。

原文稿包含公开机构标识、健康统计数据和第三方媒体。本候选题依据 PPTArena 公开数据集和 MIT 标记保留归因；正式 Release 前仍需做项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/LineGraph_TestA.pptx` | 314437 | `72fb98cc839fc54dc4b07dcdb4bf7eb8c6552ac6549d00c0010e3ab021c1f3e6` | 9 页逐页渲染、原尺寸观察、对象、文字、图表、备注和媒体审计 |
| 上游 `GroundTruth/LineGraph_GroundTruthA.pptx` | 323524 | `1bc36f7e0165538765a77b9d8ddd5df473afdca0babe95cda229daa1c286c36c` | 9 页逐页渲染、第 8 页原尺寸观察、OOXML 和图表缓存数据审计 |
| `solution/Golden/LineGraph_GroundTruthA.pptx` | 297066 | `73e05b261a46e300a0d4754123d3a4414aac12283b6bf6cd65e3a6d29035eefc` | 9 页逐页渲染、原尺寸第 8 页观察、图表数据与编辑性审计 |
| `adaptation/upstream-case.json` | 1934 | `5cce7c6e6e829aab0b0531b18896c169eda238212f9114d4663600b5c0ffdcc1` | Query、style target、路径、输入与改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| 上游 `src/evaluation_pairs_refined.json` | 165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 30 原始记录和等价答案要求 |
| 上游 `src/llm/judge.py` | 40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分调用链 |
| 上游 `agent_bench/judge_predictions.py` | 16100 | `c3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样判分和结果汇总入口 |

## 发现并修正的上游问题

冻结输入第 8 页只披露了整数百分比：

- NHOPI：18%、14%、10%、11%、11%。
- Asian：16%、11%、8%、7%、7%。
- White：12%、10%、8%、7%、7%。

上游人工 Ground Truth 的图表缓存却使用了输入没有披露的小数，例如 NHOPI 2013 年为 `0.180914996771053`，Asian 2013 年为 `0.157221594534947`。如果直接复制上游 Ground Truth，Agent 无法从输入推出这些隐藏真值。

因此本地 `build_golden.py` 从冻结输入明文重新解析数据，逐值重建图表，不沿用隐藏小数。本地 Judge 会独立从输入重算真值；上游 Ground Truth 因 15 个值不一致得分为 0，避免把不可见信息当成 Benchmark 答案。

## Golden 设计与逐页审计

输入和本地 Golden 都是 9 页、10×7.5 英寸，无宏。

- 第 8 页原数据文本框已删除；标题、副标题、注释、来源、页脚和 KFF 标识保留。
- 新图表是一张 `LINE_MARKERS` 原生 Chart，横轴为 2013–2017，纵轴固定 0%–20%，三条折线都有 7pt 圆形标记。
- NHOPI、Asian 和 White 使用三种可区分颜色；右侧图例由 3 条线、3 个标记和 3 个文本框构成，全部可编辑，且图表自带图例关闭。
- 图表位于 x=0.72、y=1.88、w=7.15、h=4.62 英寸；图例起点位于 x=8.12 英寸，与图表、标题和页脚都无重叠。
- 已逐页查看输入/Golden 的 9/9 页联系表，并原尺寸查看第 8 页。其他 8 页 LibreOffice 渲染像素相似度全为 1.0。
- `visual_qa.py` 对输入和 Golden 都报告同样 4 条 `underfilled_textbox` 通用警告，位于第 1、3、7、8 页；本题不擅自改动上游已有版式。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读取冻结输入并依任务契约独立重算。

- 候选文稿必须严格为 9 页且页面尺寸不变；ZIP 包可打开、无损坏成员和宏。
- Judge 从第 8 页冻结文本框解析年份、系列和 15 个整数百分比，再与候选图表的 OOXML 缓存和 Chart 对象比对。
- 第 8 页必须恰有一张带标记折线图；拒绝柱形图、多图表、无标记、默认图例、错误坐标轴或截图伪装。
- 自定义图例必须有 NHOPI、Asian、White 三个独立文本对象，每个对象左侧都有原生线段和标记，显式 RGB 颜色需与对应系列匹配。
- 第 8 页原有的 3 个非目标对象通过内容、样式和几何签名保护；其他 8 页通过内容、样式、图表数据、图片哈希和备注签名保护。
- LibreOffice 实际渲染冻结输入和候选文稿，非目标页最低像素相似度阈值为 0.9985；第 8 页忽略主内容可编辑区后，其余区域阈值同样为 0.9985。
- 独立合法变体从冻结输入出发，使用紫/橙/绿三色、不同图表几何和不同图例行距，不复制 Golden 坐标或颜色。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地 Golden | 1.000000 | 通过 | 13 条判据全部通过 |
| 独立颜色、几何和图例的合法替代答案 | 1.000000 | 通过 | 数据一致，不依赖单一 Golden 版式 |
| fresh replay | 1.000000 | 通过 | 从空输出目录回放，产物 SHA 与已审计 Golden 一致 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或仍为文本表 |
| 错误图表类型 / 第二张图表 / 无标记 | 0.000000 | 拒绝 | 唯一带标记折线图门槛失败 |
| 篡改任一数值 / 篡改年份 | 0.000000 | 拒绝 | 冻结输入独立重算校验失败 |
| 开启默认图例 / 缺少自定义图例 / 错误图例名 | 0.000000 | 拒绝 | 默认图例或自定义图例门槛失败 |
| 纵轴改为 5%–40% 且取消百分比格式 | 0.000000 | 拒绝 | 坐标轴语义与可读性门槛失败 |
| 保留源数据文本 / 改写第 8 页标题 / 改写其他页 | 0.000000 | 拒绝 | 目标页或非目标页保持失败 |
| 整页截图压平 / 新增第 10 页 / 夹带 `solution/Golden` | 0.000000 | 拒绝 | 可编辑性、结构或输出卫生失败 |
| 上游人工 Ground Truth | 0.000000 | 拒绝 | 15 个隐藏小数与输入明文不一致 |
| 将图表缩小到 4.8×3.2 英寸 | 0.930000 | 中等 | 内容、数据、图例和可编辑性正确，仅布局尺寸分项失败 |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena030-oracle-text-data-to-line-chart-20260829`，trial 为 `pptarena-030-text-data-to-line-c__3NzFCZL`，0 个异常；容器产物 SHA256 为 `73e05b261a46e300a0d4754123d3a4414aac12283b6bf6cd65e3a6d29035eefc`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-030-text-data-to-line-chart
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-030-text-data-to-line-chart
    harbor run --path tasks/ppt/pptarena-030-text-data-to-line-chart --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --job-name pptarena030-oracle-text-data-to-line-chart-20260829

## 剩余风险

- 正式公开 Release 前仍需对原文稿中的 Kaiser Family Foundation 标识、健康统计数据和 11 个媒体文件做项目级法务与归因复核。
- 当前已通过 OOXML、python-pptx 语义对象、LibreOffice 渲染和人工对照；正式 Release 前仍建议用 Microsoft PowerPoint 再做一次打开/保存抽检。
- 当前 L2 / Medium 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。
