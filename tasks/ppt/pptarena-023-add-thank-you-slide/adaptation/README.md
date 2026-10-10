# PPTArena Case 23 逐题适配记录

## 适配契约

- 角色：医疗研究演示文稿编辑。
- 输入：14 页 Sutter Health 医疗数据研究演示文稿。
- 真实意图：在原文稿最后追加一页以 `Thank you` 为醒目主标题的结束页。
- 主交付物：`AddThankYouSlide_GroundTruthA.pptx`。
- 不可妥协约束：新页恰好为第 15 页；`Thank you` 唯一、可编辑、醒目并合理居中；原 14 页的内容、顺序、对象、图表、媒体、备注和视觉效果不得改动。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena Case 23: Add Thank You Slide。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`，工作树干净。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入、人工 Ground Truth 和开放式布局要求已保存在 `upstream-case.json`，并与冻结索引逐字段比对。
- 上游 evaluator 使用原件、Ground Truth 和候选文件的结构化 smart diff，再交给视觉模型判断。本次完整阅读了 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，但没有执行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开演示文稿和人工 Ground Truth，不把输入、Query 或答案伪称为本项目自行编造的数据。

原文稿包含公开机构名称、人名、图表、图片和联系信息。本候选题依据 PPTArena 的公开数据集和 MIT 标记保留归因；正式 Release 前仍需重做第三方素材、人名和联系信息的法务与隐私复核。

## 已读入上下文的文件

| 文件                                           |   字节数 | SHA256                                                           | 实际检查                                             |
|------------------------------------------------|---------:|------------------------------------------------------------------|------------------------------------------------------|
| `data/input_files/AddThankYouSlide_TestA.pptx` |  1103953 | `cd6c301ffa1fe987814217e02220e63404c9481e20aebc9d2fb3c3394727ebbc` | 14 页逐页渲染、原尺寸观察、对象、图表、备注和媒体审计 |
| `solution/Golden/AddThankYouSlide_GroundTruthA.pptx` | 1111871 | `ca53da109e42215fa2e45d85cf974b56dd0e107167062e676e0fbb51e539dc1e` | 15 页逐页渲染、原尺寸观察、OOXML 和输入差异审计       |
| `adaptation/upstream-case.json`                |      891 | `1e0eea94a36fc6e55526d545bc83daa8c52e576d15f6e77b3cd41cf1cbfa259f` | Query、style target、路径与改造边界逐字段比对        |
| `adaptation/upstream-github-readme.md`          |     4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记                         |
| `adaptation/upstream-hf-dataset-card.md`        |     5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明                 |
| `adaptation/LICENSE-PPTArena-MIT.txt`            |     1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本                                      |
| 上游 `src/evaluation_pairs_refined.json`       |   165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 23 原始记录和不限定精确版式的等价答案要求         |
| 上游 `src/llm/judge.py`                        |    40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分调用链                |
| 上游 `agent_bench/judge_predictions.py`        |    16100 | `3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样判分和结果汇总入口              |

## 文稿与 Golden 逐页审计

输入为 14 页、10×7.5 英寸，共 61 个形状、52 个文字对象、96 个段落、4 张原生图表、2 个页面图片对象、3 个媒体文件和 1 组 SmartArt 数据，无隐藏页和宏。第 3–13 页包含原有非空讲者备注，其中第 12 个物理页原备注已写为 `13`，属于上游既有内容，本题不擅自“修复”。

- 人工 Golden 仅在末尾增加第 15 页；原 14 页的语义对象记录、备注、媒体和渲染像素保持。
- 第 15 页使用 Title Only 版式，含唯一原生可编辑标题 `Thank you`；标题框为 9×1.25 英寸，视觉醒目并接近页面中心。
- 结束页继承主题的青绿色左侧色带和 Sutter Health 页脚；它是上游给出的一个有效解，不是唯一允许的坐标或版式。
- 人工目视检查了输入/Golden 的 14/14 对原尺寸对照图，并单独查看了新增第 15 页。
- 既有版式问题包括：第 3 页红色圆形压住部分文字、第 6 页 `Unclassified` 换行、第 7/8 页图表紧密、部分页脚贴边。这些在输入和 Golden 中共同存在，本题不改动。
- `visual_qa.py` 对输入报告 11 条既有警告，对 Golden 报告相同 11 条加新标题框的通用高度密度警告；逐页目视未发现新增裁切、空白页或原页改动。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，并根据任务要求独立重算。

- 页数必须从 14 变为 15，页面尺寸不变，新页必须处于末尾。
- 原 14 页通过工具中立的语义签名保护文字、段落样式、对象几何、图片内容、图表数值、SmartArt 结构和备注，不依赖会被 Office 重写的原始 XML 字节。
- LibreOffice 实际渲染输入和候选文稿，原 14 页最低逐页像素相似度阈值为 0.9985。
- 第 15 页必须恰有一个完全等于 `Thank you` 的可编辑文字对象，同时检查文本框尺寸、页内位置、居中偏移和可见度。
- Judge 接受标题占位符，也接受显式字号不小于 32pt 的普通文本框；水平/垂直均用可解释的合理范围，不强制复制 Golden 坐标。
- 额外正确变体从原始 14 页出发，使用空白版式和 44pt 可编辑文本框，在不使用 Golden 布局的前提下得分 1.0。
- 整页截图、不可见白色标题、重复标题、答案路径、额外页、错误标题、原页正文改写和图表数值篡改都有专门反例。

## 实际验证结果

| 样例                              | 实际 reward | 结果 | 关键原因                                           |
|-----------------------------------|------------:|------|----------------------------------------------------|
| 上游人工 Golden                   |    1.000000 | 通过 | 11 条判据全部通过                                  |
| 独立布局的合法替代答案             |    1.000000 | 通过 | 空白版式、44pt 普通文本框，不依赖单一 Golden 布局 |
| fresh replay                      |    1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致          |
| 空交付 / 损坏 PPTX / 未修改输入   |    0.000000 | 拒绝 | 缺失、无法打开或仍只有 14 页                         |
| 标题写成 Thanks                   |    0.000000 | 拒绝 | 末页缺失精确要求的 `Thank you`                       |
| 重复添加第二个 Thank you          |    0.000000 | 拒绝 | 主标题不唯一                                      |
| 将结束页插到开头                   |    0.000000 | 拒绝 | 原页顺序和末页要求同时失败                            |
| 改写原页正文 / 篡改图表数值         |    0.000000 | 拒绝 | 冻结输入的内容或数据保护失败                           |
| 用截图压平结束页                   |    0.000000 | 拒绝 | 缺失原生可编辑标题且整页图片覆盖                       |
| 夹带 `solution/Golden` 标记         |    0.659091 | 拒绝 | 输出泄漏与未授权文字扣分，低于 0.8 通过线                 |
| 将标题设为不可见白色                |    0.670455 | 拒绝 | 末页渲染可见性失败，低于 0.8 通过线                     |
| 标题下移 1.75 英寸                  |    0.909091 | 中等 | 内容和可编辑性正确，仅合理居中分项失败                      |
| 错误增加第 16 页                    |    0.000000 | 拒绝 | 页数和末页结构变化                                  |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena023-oracle-add-thank-you-20260828`，trial 为 `pptarena-023-add-thank-you-slide__xCJAXbC`，0 个异常；容器产物 SHA256 为 `ca53da109e42215fa2e45d85cf974b56dd0e107167062e676e0fbb51e539dc1e`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-023-add-thank-you-slide
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-023-add-thank-you-slide
    harbor run --path tasks/ppt/pptarena-023-add-thank-you-slide --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 的流程使用 `render_slides.py` 渲染输入 14 页和 Golden 15 页。
- 已逐页查看原 14 页的 14/14 组原尺寸左右对照图，并单独查看第 15 页。
- 原 14 页的 LibreOffice 渲染文件逐页相似度最低值为 1.0，未出现新增裁切、离页对象、空白页或图表变化。
- 末页标题在人工 Golden 和独立合法变体中都渲染可见；白色不可见标题反例被 Judge 稳定拒绝。

## 剩余风险

- 正式公开 Release 前仍需对原文稿中的公开人名、联系方式和 3 个媒体文件做项目级隐私与法务复核。
- 上游人工 Ground Truth 的 Office 保存过程重写了部分 OOXML 和图表包装；当前已通过语义对象检查、逐页渲染和人工对照，但正式 Release 前仍建议用 Microsoft PowerPoint 再做一次打开/保存抽检。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 任务专属视觉评分补充

制题模型重新完整读取了 instruction、task.toml、deterministic rubric、完整 judge、expectations、solution、upstream case 和适配记录，并将 14 页输入与 15 页 Golden 分别渲染为联系图；逐页检查了封面、地图、饼图、数据结论、Acknowledgements 和新增第 15 页的原尺寸渲染。视觉评分采用 `reference_free` 模式，因为上游明确允许使用与原文稿一致或其他简洁专业、合理居中的结束页布局；Golden 只用于理解可接受完成度，不强迫候选复制其精确坐标。

`visual_rubric.json` 为本题单独定义五个维度：前 14 页研究文稿的可见保持、第 15 页 `Thank you` 的焦点与合理居中、末页与 Sutter Health 青绿品牌语言的自然收束、标题投影可读性与无关文本控制，以及无截图/遮罩/调试痕迹。每项均写明 0、0.5、1 的可见锚点和必须检查的页面证据。

Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、the configured VLM provider 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。

中等可见负例保留了原 14 页的青绿色带、地图、饼图、研究文本、Sutter Health 页脚和页序，新增第 15 页也仍有唯一、清晰、原生可编辑的 `Thank you`，但标题被移到明显偏右下的位置，破坏了结束页的几何中心和留白平衡。其代码分为 0.579545，the VLM evaluator 视觉分为 0.860000，按 70/30 权重得到综合分 0.663682；视觉分严格低于 Golden，且高于空白或损坏文件。本题 6 次 the VLM evaluator 调用均一次成功，未发生基础设施失败；任何后续 API、渲染或解析失败只记为评测基础设施失败，不计作 0 分。
