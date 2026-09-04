# PPTArena Case 37 逐题适配记录

## 适配契约

- 角色：漫画课程教学演示文稿编辑。
- 输入：9 页《Panel to Panel Transitions in Comics》教学演示文稿。
- 真实意图：为全部 9 页添加幻灯片级 `Fade` 转场。
- 主交付物：`AddTransitions_GroundTruthA.pptx`。
- 不可妥协约束：9/9 页都是 Fade；无自动换页、声音或其他转场；文字、图片、布局、样式、顺序、备注和动画均不改动。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 37: Add Transitions。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入、人工 Ground Truth 和宽松速度/变体要求已保存在 `upstream-case.json`，并与冻结索引逐字段比对。
- 上游 evaluator 通过 smart JSON diff 和视觉模型判断候选文稿，明确 Fade 的速度、时长和 Through Black 变体不锁定。本次完整阅读了 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，但未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开演示文稿和人工 Ground Truth，不把输入、Query 或答案伪称为本项目自行编造的数据。

必须区分“PPTArena 数据集代码/数据卡标记 MIT”和“演示文稿内部第三方漫画页面的版权”。该文稿包含 Trung Le Nguyen 的《The Magic Fish》图像；本题在候选阶段保留上游归因和风险标记，正式公开 Release 前必须单独完成版权法务审查或替换授权素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/AddTransitions_TestA.pptx` | 945680 | `a5557a252a8b8668a916a106b54c9935dd5d1eef2eb76d5f2fa4efa5bd4529b4` | 9 页逐页渲染和原尺寸观察；对象、图片、备注、转场、动画和包部件审计 |
| `solution/Golden/AddTransitions_GroundTruthA.pptx` | 941359 | `c9b75e9063f87e8aac5c262c4b32a48d572119a130c50a0fd04e52108ff209db` | 9 页逐页渲染和原尺寸观察；逐页 OOXML Fade 和输入差异审计 |
| `adaptation/upstream-case.json` | 838 | `e870eca718b129ce48271403babb4ff9e904f5aaa3f3dd0d1dda9dd643a42a9b` | Query、style target、路径与改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| 上游 `src/evaluation_pairs_refined.json` | 165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 37 原始记录和不锁定转场时长/变体的要求 |
| 上游 `src/llm/judge.py` | 40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分调用链 |
| 上游 `agent_bench/judge_predictions.py` | 16100 | `3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样判分和结果汇总入口 |

## 文稿与 Golden 逐页审计

输入和 Golden 都是 9 页、10×5.625 英寸，共 27 个形状、19 个文字对象、8 个图片对象和 8 个媒体文件，无图表和宏。第 2–9 页含长篇教学备注，第 1 页无实质备注。

- 冻结输入 9/9 页都无 `<p:transition>`；人工 Golden 9/9 页都是 `<p:transition spd="slow"><p:fade/></p:transition>`。
- 输入和 Golden 的 ZIP 成员集都为 88 个部件；只有 `docProps/app.xml`、`docProps/core.xml` 和 9 个幻灯片 XML 字节发生变化。
- 从 9 个 Golden 幻灯片 XML 删除 `<p:transition>` 后，每页 `<p:cSld>` 的规范化 XML 哈希均与输入完全一致。
- 输入和 Golden 的 8 个媒体哈希、页面对象签名、版式分配、讲者备注和动画/时序签名全部一致。
- 已原尺寸目视检查输入/Golden 的 9/9 组左右对照图；静态渲染逐页相似度全部为 1.0，符合“只添加播放时转场”的任务语义。
- `visual_qa.py` 对输入和 Golden 都报告相同的 3 条 `underfilled_textbox` 通用警告，位于第 4、5、6 页；它们是上游既有版式，不属于本题的转场编辑范围。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，并根据任务契约独立重算。

- 逐个解析 `ppt/slides/slide1.xml`–`slide9.xml`；每页必须恰有一个幻灯片级 `p:transition`，直接效果必须是唯一 `p:fade`。
- 允许未设置速度，也允许 `slow`、`med`、`fast`；允许合理时长和 Through Black 布尔变体。
- 明确拒绝 `advTm` 自动换页、`advClick=0`、转场声音、Cut/其他效果、重复转场节点和只改母版。
- 逐页对比文字、run 样式、占位符、对象顺序、几何、旋转、图片哈希/裁剪、隐藏状态和版式分配。
- 备注对象和文字逐页保护；`p:timing` 节点独立保护，用于拒绝新增、删除或改写动画/时序。
- 媒体哈希 Counter、包成员集和整页图片覆盖独立检查，防止丢媒体、夹带新媒体和截图压平。
- LibreOffice 实际渲染冻结输入和候选文稿，9 页最低逐页像素相似度阈值为 0.9985。
- 字节级保持只是 5 分正向分项，不是锁死某一编辑器的核心门槛；无害重打包可以在语义、视觉和转场全部正确时得到 0.95。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 上游人工 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立默认 Fade 答案 | 1.000000 | 通过 | 全部 9 页不指定速度，不复制 Golden 的 `slow` |
| 独立 Fast + Through Black 答案 | 1.000000 | 通过 | 合理变体被接受，说明 Judge 未锁死单一转场参数 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.950000 | 中等 | 语义、转场、对象和渲染正确，仅字节级保持分项失败 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或 9 页仍无转场 |
| 漏 1 页 / 1 页 Cut / 全部 Cut / 重复转场 | 0.000000 | 拒绝 | 9/9 页唯一 Fade 门槛失败 |
| 自动换页 / 禁用点击换页 / 转场声音 | 0.000000 | 拒绝 | 夹带未授权播放行为 |
| 改写正文 / 改写备注 / 新增动画 | 0.000000 | 拒绝 | 内容、备注或时序保护失败 |
| 篡改媒体 / 用整页图片压平第 1 页 | 0.000000 | 拒绝 | 媒体哈希、原生对象和可编辑性失败 |
| 夹带 `solution/Golden` 标记 | 0.000000 | 拒绝 | 未授权文字、对象保持和渲染保持同时失败 |
| 错误新增第 10 页 | 0.000000 | 拒绝 | 页数、顺序和包结构变化 |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena037-oracle-fade-transitions-20260828`，trial 为 `pptarena-037-fade-transitions-al__YVkCyRt`，0 个异常；容器产物 SHA256 为 `c9b75e9063f87e8aac5c262c4b32a48d572119a130c50a0fd04e52108ff209db`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-037-fade-transitions-all-slides
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-037-fade-transitions-all-slides
    harbor run --path tasks/ppt/pptarena-037-fade-transitions-all-slides --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 的流程使用 `render_slides.py` 渲染输入、人工 Golden 和 Fast + Through Black 独立合法变体。
- 已逐页查看输入/Golden 的 9/9 组原尺寸对照图，并原尺寸复查独立合法变体的第 1 页和第 9 页。
- 三份文稿均可由 LibreOffice 渲染为完整 9 页；候选文稿与冻结输入的 9 页最低静态像素相似度为 1.0。
- 转场是播放时行为，PDF 静态渲染不能展示淡入过程；因此转场真值由逐页 OOXML 结构检查覆盖，静态渲染专门用于确认未改动正文和版式。

## 剩余风险

- 正式公开 Release 前必须重点处理《The Magic Fish》漫画页面的第三方版权；上游数据卡的 MIT 标记不能自动消除该风险。
- 当前已通过 OOXML、python-pptx 语义对象、LibreOffice 渲染和人工对照；正式 Release 前仍建议用 Microsoft PowerPoint 实际播放一次，抽检 Fade 动态效果与 Through Black 变体。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、任务专属 `solve.py`、冻结输入和人工 Golden，核对了 9 页每页唯一 `p:fade`、转场安全变体、原页对象/备注/动画签名与 8 个媒体。
- 冻结输入和 Golden 均已完整渲染为 9 页、制作联系图并逐页亲眼对照；两者的封面、六种漫画转场概念/案例和 Group Activity 页在静态渲染中一致，符合“只增加播放时 Fade”的任务边界。
- 由于任务不允许任何静态可见改动，本题使用 `reference`；专属视觉 Rubric 逐页检查静态保持、教学文本层级、漫画取景/清晰度、白蓝黑 + 暖色漫画的系列连续性和无截图/遮罩/调试痕迹的交付干净度。
- Fade 是 PDF 联系图无法显示的播放时行为，9/9 页转场类型、速度/时长/通过黑色变体、自动换页/声音禁用、备注/动画、对象几何、媒体哈希和原生可编辑性继续由确定性 Judge 负责；VLM 不根据静态图猜测 Fade，只评定是否产生任务外的可见回归。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次综合分均为 `1.0000`。
- 中等可见负例从 Golden 手工定向改造：完整保留 9/9 页 Fade、全部转场参数、讲者备注、文字、图片、媒体与原生对象，仅把第 9 页 Group Activity 漫画放大并左移，使图片明显遮挡问题列表的右侧内容。已完整渲染并亲眼确认：前 8 页完好，最后一页仍能辨认标题、问题与漫画，但存在肉眼明显的图文重叠。因本题确定性 Judge 对任务外静态改动采用严格门槛，代码分为 `0.0000`；视觉分 `0.5000`、综合分 `0.1500`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮部分调用触发了 Provider 长尾和配置内自动重试，最终全部成功并写入完整结果；无未解决基础设施失败，未将任何失败冒充为 0 分。
