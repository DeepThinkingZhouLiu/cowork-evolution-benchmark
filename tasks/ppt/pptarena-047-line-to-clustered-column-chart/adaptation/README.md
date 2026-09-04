# PPTArena Case 47 逐题适配记录

## 适配契约

- 角色：演示文稿图表编辑。
- 输入：1 页 4:3 美国成人识字率比较幻灯片，内含唯一原生折线图。
- 真实意图：只将折线图改为簇状柱形图，保留底层数据、图表位置、数据标签和其他页面对象。
- 主交付物：`ChangeLineGraphtoBar_GroundTruthA.pptx`。
- 不可妥协约束：原生可编辑的簇状柱形图；3 个分类和精确浮点数值不变；无图例/图表标题；标题、页码占位符、显著性注释、Logo 和备注不变。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 47: Change Line Graph to Bar Graph。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>，commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可文本与数据卡副本保存在本目录。
- 上游 Query、style target、enhancement note、输入与人工 Ground Truth 路径已逐字段保存在 `upstream-case.json`。
- 已阅读 `src/evaluation_pairs_refined.json`、`src/llm/judge.py`、`agent_bench/judge_predictions.py` 和公开 Agent 结果；未运行需要外部模型的网络 Judge。
- 本题沿用上游真实 PPTX、原生图表、教育统计文本和 AIR Logo，不是本项目编造的 Mock 输入。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/ChangeLineGraphtoBar_TestA.pptx` | 58165 | `badd9134a75308729132505dc232413229bf20914fa0fddb42f4bf186b420114` | 1 页渲染；5 个页面对象、图表类型、数据缓存、坐标轴、几何、备注和包成员审计 |
| 上游 `GroundTruth/ChangeLineGraphtoBar_GroundTruthA.pptx` | 54069 | `0a9290cf301ba657d3b1eace9a06374c77625cefd7f05424fb30bd24f24681ce` | 1 页渲染；柱形图、数据、页码差异和整包差异审计 |
| `solution/Golden/ChangeLineGraphtoBar_GroundTruthA.pptx` | 58161 | `6ae3c3b46082ec60ac053015c0e7f3761531a0ba866c00a2d82015f480e30873` | 仅图表部件改动；原生可编辑结构、视觉和非图表区域审计 |
| `adaptation/upstream-case.json` | 1093 | `a1eec201b1c708be2f88ac45f7338f5b18f2f5f673319969b537f4ed1abb330a` | Query、style target、enhancement note、路径和适配边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 仓库结构、人工 Ground Truth 和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、数据介绍和公开状态 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| `tests/fixtures/golden-slide1.png` | 94539 | `87811dd6ad711e48717ac1fcaa43fd2552c18aeca63de33bf005737b1a50afa7` | 只用于构造整页压平负例 |

## 图表真值与 Golden 构建

- 冻结输入图表为 `LINE_MARKERS`，单系列，无图例、无图表标题，数值轴范围 245–285，主单位 10。
- 三个分类为 `IALS 1994`、`ALL 2003`、`PIAAC 2012/2014`；精确缓存值为 273.475674150447、268.246769360089、271.722770340152。
- 第二个数据点的自定义标签包含显著性星号，页脚原文为 `*p < .05. Average score is significantly different from PIAAC.`。
- `build_golden.py` 仅接受锁定 SHA256 的本题输入，把 `lineChart` 改为 `barChart`，设置垂直簇状分组、柱间距和可编辑填充，并移除折线专用 marker/smooth 元素。
- 本地 Golden 与冻结输入的 ZIP 成员列表完全一致，只有 `ppt/charts/chart1.xml` 的字节内容不同；标题、页码占位符文本 `14`、底部注释、Logo、备注和包结构均保留。
- 独立合法变体将柱间距改为 150、填充改为主题色 2，并在容差内微调几何；它不复制 Golden 的样式参数，依然得分 1.0。

## 上游 Ground Truth 审计

- 上游 Ground Truth 的柱形图类型、3 个分类、数值、数据标签、图例状态和图表几何均正确。
- 它不是最小编辑：上游将原包中的 `slide14.xml` / `notesSlide10.xml` 重写为 `slide1.xml` / `notesSlide1.xml`，并改写大量 master、layout、relationship 和文档属性。
- 更重要的是，上游将页码占位符底层文本从 `14` 改成 `1`；这与“其他元素不变”相冲突。因此本地 Golden 不直接复制上游文件。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只使用冻结输入、上游任务契约和本题可公开重算的真值。

- 要求唯一原生 Chart 对象的类型为 `COLUMN_CLUSTERED`；水平条形、堆积柱形、原折线和整页图片均失败。
- 通过图表缓存独立重算分类顺序、单系列名和 3 个精确数值，不认可只在页面上写出相似整数。
- 逐对象比较全部非图表对象的类型、占位符类型、文本、段落/run 字体、几何和图片 blob，同时比较备注与媒体集。
- 检查无图例、无图表标题且有可见数据标签；图表位置/尺寸每轴容差为 0.25 英寸。
- 使用独立 LibreOffice profile 渲染输入与候选文件；屏蔽图表矩形后比较非图表像素，并检查图表区域有足够非白像素。
- 坐标轴 245–285 范围与 10 的主单位是正分，避免把可读性轻微差异错当成内容严重错误。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地 Golden | 1.000000 | 通过 | 10 条判据全部通过 |
| 独立红色窄间距合法变体 | 1.000000 | 通过 | 类型、数据、标签、保持与几何合格 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物与已审计 Golden 哈希一致 |
| 空交付 / 损坏 PPTX / 未修改折线图 | 0.000000 | 拒绝 | 缺文件、无法打开或图表类型错误 |
| 水平条形 / 堆积柱形 / 错数值 / 错分类 | 0.000000 | 拒绝 | 类型或底层数据门槛失败 |
| 删数据标签 / 加图例 / 过度移动或缩小 | 0.000000 | 拒绝 | 标注与布局契约失败 |
| 页码 14 改为 1 / 改标题 / 删 Logo | 0.000000 | 拒绝 | 非图表对象保持门槛失败 |
| 删图表 / 额外图表 / 整页截图 | 0.000000 | 拒绝 | 唯一原生可编辑图表门槛失败 |
| 夹带 `solution/Golden` | 0.000000 | 拒绝 | 非图表对象改动且命中答案标记 |
| 只将坐标轴下限改为 240 | 0.930000 | 中等 | 核心真值正确，仅坐标轴可读性正分失败 |

同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena047-oracle-line-to-column-20260829`，trial 为 `pptarena-047-line-to-clustered-c__VtNpL7f`，0 个异常；容器产物 SHA256 为 `6ae3c3b46082ec60ac053015c0e7f3761531a0ba866c00a2d82015f480e30873`，与已审计 Golden 一致。

## 视觉检查与剩余风险

- 冻结输入、上游 Ground Truth、本地 Golden 和独立合法变体均使用独立 LibreOffice profile 按原尺寸渲染并人工查看。
- `visual_qa.py` 对冻结输入、本地 Golden 和独立合法变体均报告 0 条启发式警告。
- LibreOffice 渲染动态页码字段时显示当前文件中的第 1 页，而 OOXML 页码占位符底层文本为 `14`；Judge 比较的是未被 LibreOffice 重算前的原生占位符。
- 正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检；L2 / Medium 仅是制题配比目标，正式难度必须经过冻结 Baseline Agent 重复 Rollout 后再写回。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、本题 `build_golden.py` / `build_fixtures.py` / `solve.py`、冻结输入和本地 Golden，核对了唯一原生图表、三类调查、三个精确缓存值、自定义 `268*` 标签、245–285 轴范围、图表几何与全部非图表保持规则。
- 冻结输入与本地 Golden 均已原尺寸渲染并亲眼对照：输入以蓝色折线和菱形点呈现 273、268*、272；Golden 在同一主体区域改为三根等宽蓝色竖向柱，保留三类标签、五条主刻度、灰底结论标题、AIR Logo、显著性注释和页码。
- 本题是有明确数据和目标图表形态的视觉编辑，因此使用 `reference`；允许不复制 Golden 的主题色号或精确柱间距，但应保持统计关系和原页信息层级。专属视觉 Rubric 分别检查簇状柱形语义、273/268*/272 与 245–285 轴的真实可读性、克制的统计图样式、标题—图表—页脚安全区和非图表区域的整体交付质量。
- 图表原生类型、三个精确浮点数、分类缓存、数据标签 XML、无图例/标题、几何容差、非图表对象签名、备注和包完整性继续由确定性 Judge 负责；VLM 只补充三柱感知清晰度、轴范围是否视觉误导、柱色/网格/标签间距、留白和专业完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次综合分均为 `1.0000`。
- 中等可见负例直接使用本题已有的 `axis-changed` 边界样例：三根蓝柱、273/268*/272 标签、三个分类、灰底标题、AIR Logo、页脚和原生可编辑性均保留，但纵轴由 245–285 改成 240–280，刻度变为 240/250/260/270/280，柱体从错误的更低基线起算，视觉上放大调查间差异。已原尺寸渲染并亲眼确认文件可打开、内容可读且缺陷肉眼可见；该负例确定性分 `0.9300`、视觉分 `0.8800`、综合分 `0.9150`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮 6 次 VLM 调用均成功完成，没有基础设施失败或被错误计作 0 分的样本。
