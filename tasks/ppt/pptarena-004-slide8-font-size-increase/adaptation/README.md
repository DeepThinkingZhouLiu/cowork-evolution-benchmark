# PPTArena Case 4 逐题适配记录

## 适配契约

- 角色：教学演示文稿编辑。
- 输入：56 页凹面镜与光学教学演示文稿。
- 真实意图：只在第 8 页把指定正文精确增大 2 pt，同时保护标题、图中标签、页码和整份文稿。
- 主交付物：`FontSizePlus2Slide8EditedA.pptx`。
- 不可妥协约束：九个目标对象字号正确，例外对象不变，非目标页不变，内容仍为原生可编辑 PPTX。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena `Case 4: Increase Font Size on Slide 8`。
- GitHub：`https://github.com/michaelofengenden/PPTArena.git`。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在 `upstream-github-readme.md`、`upstream-hf-dataset-card.md` 和 `LICENSE-PPTArena-MIT.txt`。
- 上游 Query、style target、输入和人工 Ground Truth 映射完整保存在 `upstream-case.json`；该 JSON 已与冻结索引逐字段比对，语义完全一致。
- 上游数据卡明确称 Ground Truth 为人工编辑目标，本题不把输入或答案伪称为自行编造的数据。

原文稿为公开数据集中的教学材料。本候选题依据上游 MIT 数据集授权保留完整归因；在正式 Release 前仍需纳入项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件                                                     | 字节数 | SHA256                                                           | 实际检查                                             |
|----------------------------------------------------------|-------:|------------------------------------------------------------------|------------------------------------------------------|
| `data/input_files/FontSizePlus2Slide8TestA.pptx`         | 1012979 | `6beba374b051d77ba11636d826535fc3e391f245daf14d7547383a02798affd6` | 56 页渲染、联系表、第 8 页对象、文字运行和几何展开检查 |
| `solution/Golden/FontSizePlus2Slide8EditedA.pptx`        | 1007042 | `0dcf73c3179e7a98438362eb96b2e2b84af00fedfea1c2e7071f36cf4d3f45f1` | 56 页渲染、逐页原尺寸查看、OOXML 差异和字体属性检查    |
| `upstream-case.json`                                     |    1770 | `b3631f30dccc8414578ad1265795373e62a631f4768e94ecdcbd25948af77849` | Query、style target 与索引逐字段比对                  |
| `upstream-github-readme.md`                              |    4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记                     |
| `upstream-hf-dataset-card.md`                            |    5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态、人工 Ground Truth 说明             |
| `LICENSE-PPTArena-MIT.txt`                               |    1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本                                      |

## 上游答案审计与适配决策

上游输入和 Golden 均为 56 页、13.333×7.5 英寸，共有 445 个文本元素。逐对象 OOXML 差异确认只有第 8 页九个文本对象发生了字号变化：

| 目标文字                                                    | 原字号              | Golden 字号          | 其他变化                       |
|-------------------------------------------------------------|----------------------|----------------------|--------------------------------|
| `Object height (h)`、`Image height (h')`                    | 24 pt                | 26 pt                | 无                             |
| `Object distance (s)`、`Image distance (s')`                | 24 pt                | 26 pt                | 无                             |
| `For an object outside ...`、`Focal length ...`             | 20 pt                | 22 pt                | 无                             |
| `Can be reduced or enlarged`                                | 24 pt                | 26 pt                | 无                             |
| `Inverted` / `(upside down)`                                | 24 pt / 继承约 18 pt | 26 pt / 20 pt        | 文本框高增加约 0.067 英寸      |
| `Real` / `(formed by intersection of actual rays)`          | 24 pt / 继承约 18 pt | 26 pt / 20 pt        | 文本框高增加约 0.101 英寸      |

主标题占位符、六个斜体单字符标签 `C`、`F`、`h`、`h'`、`s`、`s'` 和自动页码均保持不变。九个目标对象的位置、宽度、字体、颜色、粗斜体和对齐均保持；只有两个启用自动适应的文本框自然增高。

除第 8 页外的 55 页渲染文件逐字节一致。PowerPoint 保存时只在第 8 页内容之外重写了文档属性、视图属性和 custom XML 元数据，因此本题不做整包哈希相等判断，避免把合法 Office 重保存误判为错误。

本地 Judge 不读取 Golden，也不比较 Golden 哈希。它从冻结输入独立重算九个目标对象、逐运行有效字号、非字号文字样式、例外对象、56 页文字、图片媒体、第 8 页非文本图形、55 个非目标页渲染、PDF 第 8 页可提取文字、页面边界和输出目录卫生。

## 实际验证结果

| 样例                               | 实际 reward | 结果 | 关键原因                                                |
|------------------------------------|------------:|------|---------------------------------------------------------|
| 上游人工 Golden                    |    1.000000 | 通过 | 11 条判据全部通过                                      |
| fresh replay                       |    1.000000 | 通过 | 从空输出目录重放，产物 SHA 与人工 Golden 一致          |
| 空交付                             |    0.000000 | 拒绝 | 缺失指定交付物                                         |
| 损坏 PPTX                          |    0.000000 | 拒绝 | ZIP / OOXML 无法打开                                  |
| 未修改输入                         |    0.000000 | 拒绝 | 九个目标对象没有增加 2 pt                              |
| 所有目标只增加 1 pt                |    0.000000 | 拒绝 | 精确字号规则失败                                       |
| 只修改一个目标对象                 |    0.000000 | 拒绝 | 八个目标未完成                                         |
| 误改主标题                         |    0.000000 | 拒绝 | 非目标例外样式变化                                     |
| 误改六个光线图标签                 |    0.000000 | 拒绝 | 非目标例外字号变化                                     |
| 改写第 1 页文字                    |    0.000000 | 拒绝 | 原文和非目标页保护失败                                 |
| 将目标文字改为红色                 |    0.000000 | 拒绝 | 非字号样式保护失败                                     |
| 第 8 页整页截图                    |    0.000000 | 拒绝 | 原生文本缺失、媒体集变化并新增全页图                   |
| 重复一个目标文本框                 |    0.000000 | 拒绝 | 对象唯一性、原文和样式检查失败                         |
| 夹带 `solution/Golden` 标记        |    0.000000 | 拒绝 | 新增未授权文字与答案路径                               |
| 字号正确但把目标框移出页面         |    0.777778 | 中等 | 核心内容正确，但几何和可读性交付失败                   |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 `1.0`，总体标准差 `0.0`。

官方 Harbor 0.21 Oracle 容器最终回放结果为 `1.0`，trial 为 `pptarena-004-slide8-font-size-in__v6Hy6UK`，0 个异常，artifact SHA256 为 `0dcf73c3179e7a98438362eb96b2e2b84af00fedfea1c2e7071f36cf4d3f45f1`。

首轮 Harbor 回放为 `0.888889`：容器缺少 Calibri 的度量兼容字体，LibreOffice 回退到 DejaVu Sans 后真实裁切了底部两段文字。环境加入开源 Carlito / Caladea 字体包后，在同一容器中重新渲染并目视确认文字完整，第二轮 Harbor 回放达到 `1.0`。本题没有通过删除可读性判据来掩盖环境问题。

主要命令：

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-004-slide8-font-size-increase
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-004-slide8-font-size-increase
harbor run --path tasks/ppt/pptarena-004-slide8-font-size-increase --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes
```

## 视觉复核

- 已查看原输入和 Golden 的 56 页联系表。
- 已逐页查看 Golden 的 56/56 页原尺寸渲染图。
- 已在最终 Harbor 同构 Docker 字体环境中重新渲染并查看第 8 页；底部三组说明、图中标签和页码均完整可读。
- 除第 8 页外的 55 页没有新增占位符、空白页、裁切、离页或新的遮挡。
- 原输入中少数页面本来就有高密度排版和边缘贴近，本题要求保持；这些既有特征不计为候选答案新增缺陷。

## 剩余风险

- 正式公开 Release 前仍需对上游教学材料中的第三方内容做项目级法务复核；当前依据是 PPTArena 公开数据卡和仓库的 MIT 标记。
- 当前 `L1 / Easy` 仅是制题配比目标，不是正式难度标签；必须经过真实 Agent Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配与校准

制题模型已逐项读取 `instruction.md`、`task.toml`、`rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、上游 Case 记录、完整解题脚本、原始输入和人工 Golden。输入与 Golden 均已渲染 56 页联系图并逐页对照，另外以原始分辨率检查第 8 页的标题、九个目标文本、灰色说明条、底部三组结论和凹面镜光线图。

本题是精确字号编辑与整稿保真任务，因此采用 `reference` 模式。专属视觉 Rubric 共 5 条：56 页非目标保真、九组增大文本的容纳/换行、文字与光线图的防遮挡、例外文字与目标文字的层级/样式，以及整稿原生可编辑观感。代码 Judge 仍独立检查精确 +2 pt、例外对象、OOXML 样式、几何、媒体和 55 页渲染保真；VLM 不重复冒充结构检查，只补充实际可见的可读性、平衡、层级和专业完成度。综合分按代码分 70% + VLM 分 30% 计算，两个分项单独保存。

固定 Codex CLI 0.148.0、the configured VLM provider 对 Golden 的 5 次独立真实视觉评分均为 1.000，均值 1.000、总体标准差 0.000（低于 0.05）。中等可见负例仍保留 56 页且可正常打开，只在第 8 页把本应精确增大 2pt 的若干目标文字进一步放大并偏移，造成右上文字相互挤压、底部三组说明贴边且局部截断，同时不破坏其他 55 页。该负例视觉分为 0.685，严格低于 Golden 且高于空白/损坏交付；确定性分 0.000，按 70%/30% 合成为 0.2055。六次有效评分均成功，未把任何失败调用计为 0。
