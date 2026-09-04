# PPTArena Case 28 逐题适配记录

## 适配契约

- 角色：演示文稿编辑。
- 输入：1 份 6 页、16:9 的 PPTX，页面与议程顺序被打乱，且 `Engaging the audience` 页缺少页码。
- 真实意图：重排原有页面，使议程、内容流程和页码一致，同时保护所有未要求修改的内容与样式。
- 主交付物：`FixAgendaAndOrder_GroundTruthA.pptx`。
- 不可妥协约束：页序、议程、页码正确；标题、正文、日期、图片、来源、备注、关系、主题、母版和版式保持；全部页面保持原生可编辑。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 28: Fix Slide Order and Agenda。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>；commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本已保存。
- 上游 Query、style target、页面顺序、页码缺口、版本与适配边界逐字段保存在 `upstream-case.json`。

文稿使用了模板和 32 个媒体部件。PPTArena 的 MIT 标记不会自动消除模板和摄影图片的第三方权利风险，正式 Release 前必须再做权利复核或等价换图。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/FixAgendaAndOrder_TestA.pptx` | 7640124 | `bcbdf13b305029f1fca7c33d0b6b0020f84acf963bf066471d4be00971c2175b` | 6 页逐页原尺寸渲染；标题、正文、日期、页码、图片、链接、页面关系、32 个媒体和 176 个包成员盘点 |
| `solution/Golden/FixAgendaAndOrder_GroundTruthA.pptx` | 7635650 | `6079f8954c888ff2b10cc34e8bb6acfee0895b474edb10a4018739f4e018c554` | 从冻结输入重建；全部 6 页原尺寸目视、OOXML 差异、资产哈希和 LibreOffice 渲染审计 |
| `adaptation/upstream-ground-truth-package-mutated.pptx` | 7638710 | `0ee793bcbff7eaf001ee8727dd65358bbfac8713fbdea4eca1b23bde0ea9149a` | 上游人工 Ground Truth；全部 6 页目视，并确认漏页码、改媒体和重写关系的副作用 |
| `adaptation/build_golden.py` | 5994 | `589a5a2f6a02736764fddddd86ae91e7bd8eda40cbbb204c59989818c8b7153b` | 任务专属重建脚本；冻结输入哈希、六个关系 ID、五个议程段落和页码映射显式写入 |
| `adaptation/upstream-case.json` | 3317 | `fc5964da3cdaed54f5c0525155f08731883560dda665a2d08109cd923e8c4c5d` | Query、style target、来源版本、输入顺序、上游缺陷和适配边界逐项比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入与上游 Ground Truth 审计

冻结输入的原顺序为 `Why visuals matter -> Agenda -> Conclusion -> Engaging the audience -> Q&A -> Effective delivery techniques`。Agenda 的五项则为 `Conclusion -> Q&A -> Visual appeal -> Engaging the audience -> Effective delivery techniques`。原页面号是 1、2、3、缺失、5、6；第 4 个原页面的版式已有页码占位符定义，但实例页未放入页码对象。

上游 Ground Truth 的可用证据是逻辑顺序和议程顺序：`Agenda -> Why visuals matter -> Engaging the audience -> Effective delivery techniques -> Conclusion -> Q&A`。但它仍漏掉第 3 页页码，而且相对冻结输入无关改写 21 个共有包部件：6 个 slide XML、6 个 slide relationship、presentation、notes relation、3 个文档属性/缩略图以及 4 个媒体。它还删除 `image31.jpg` / `image32.jpeg` 并新增交换扩展名的 `image31.jpeg` / `image32.jpg`。因此上游文件只作设计证据和负例，不直接复用为本地 Golden。

## 本地 Golden 构造

本地 Golden 不是通用批量 mock，而是在逐页阅读标题、正文、页脚、图片和上游答案副作用后，从冻结输入单题重建：

- 只在 `ppt/presentation.xml` 中移动原有 `sldId` 节点，关系 ID 依次为 `rId3/rId2/rId5/rId7/rId4/rId6`。
- 在原 `slide2.xml` 中移动五个已有议程段落，不重写其文本样式。
- 将六个页码更新为 1–6；对原 `slide4.xml` 从兼容的页面复制原生 `sldNum` 占位符，仅补页码 3。
- 176 个包成员的集合完全不变；实际变化仅为 `presentation.xml` 和 6 个 slide XML。32 个媒体、六个 slide relationship、主题、母版、版式、备注和文档其他部件字节保持。

## Judge 设计

本地 Judge 不读取 Golden，也不比较 Golden 文件哈希；只从候选 PPTX 和冻结输入独立重算。

- 从六页可见标题独立判定页面身份和逻辑顺序，不依赖当前 slide 文件名。
- 逐段落提取 Agenda 五项，检查内容集、重复和顺序。
- 逐页提取原生 `sldNum` 占位符，检查 1–6、对象数、右侧页脚几何与方向；不依赖 Golden 中的对象 ID。
- 按页面身份对比对象类型、占位符类型、名称、几何、文本、段落/字体样式、图片哈希与裁剪；规范化掉允许改动的议程文本和页码。
- 比对备注、媒体哈希 Counter、包成员集、滑页关系、主题、母版和版式，专门拒绝上游 Ground Truth 的媒体与关系重写。
- LibreOffice 实际渲染冻结输入与候选文稿；按页面身份对齐，遮罩议程和页码区域后，六页像素相似度均不得低于 0.998。
- 检查对象数为 `5/5/5/5/9/5`、无全页截图压平、无额外页、输出目录干净且无答案标记。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地逐页重建 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 只替换新页码对象的合法 ID | 1.000000 | 通过 | 字节不同，但语义、几何、可编辑性和渲染均正确 |
| Agenda 条目增加无害首尾空格 | 1.000000 | 通过 | 规范化文本后语义和视觉不变 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.980000 | 中等 | 功能、页序、议程、页码和渲染正确，仅 2 分包骨架字节保持失败 |
| 只把第 1 页页码改为 9 | 0.620000 | 中等但不通过 | 内容与页序正确，但明确页码要求失败并触发严重错误扣分 |
| 空交付 / 损坏 PPTX / 未修改原件 | 0.000000 | 拒绝 | 缺失、无法打开或页序/议程/页码未完成 |
| 上游 Ground Truth | 0.000000 | 拒绝 | 漏页码 3，且改写媒体、关系和页面包结构 |
| 错页序 / 错议程 / 重复议程项 / 缺页码 3 | 0.000000 | 拒绝 | 结构、议程或可编辑性门槛失败 |
| 改标题 / 改正文 / 改日期 | 0.000000 | 拒绝 | 未要求内容保持失败 |
| 改媒体 / 改主题 / 改版式 / 改关系 | 0.000000 | 拒绝 | 文稿资产和全局结构保持失败 |
| 额外对象 / 额外页 / 答案标记 | 0.000000 | 拒绝 | 对象集、页数或交付卫生失败 |

`expectations.json` 共记录 24 组 Golden、独立正例、fresh replay、中等样例和负例，实际结果全部与期望一致。同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。负例 PPTX 由 `tests/build_fixtures.py` 可重建，验证后已删除不提交，仅保留构造器、期望表和说明。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena028-oracle-logical-slide-order-and-agenda-20260828`，trial 为 `pptarena-028-logical-slide-order__b87rwYM`，0 个异常；容器产物 SHA256 为 `6079f8954c888ff2b10cc34e8bb6acfee0895b474edb10a4018739f4e018c554`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-028-logical-slide-order-and-agenda
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-028-logical-slide-order-and-agenda
    harbor run --path tasks/ppt/pptarena-028-logical-slide-order-and-agenda --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --jobs-dir /tmp/cowork-harbor-jobs --job-name pptarena028-oracle-logical-slide-order-and-agenda-20260828

## 视觉复核

- 已按 presentation-skill 流程渲染并原尺寸查看冻结输入、上游 Ground Truth 和本地 Golden 的全部 6 页。
- 已原尺寸查看两个独立合法变体的受影响页；Agenda 五项完整可读，页码 3 与现有右侧页脚一致。
- `visual_qa.py` 对本地 Golden 只报告 7 个原文稿已存在的标题/署名文本框过高启发式警告；这些对象未被修改。
- Judge 遮罩议程和页码区域后，本地 Golden 六页与对应原页的逐页渲染相似度均为 1.0。

## 剩余风险

- 正式公开 Release 前必须复核页面模板和 32 个媒体部件的第三方权利；如无明确授权，需替换为可公开素材并保留相同的页序和议程错位结构。
- 结论页三条 arXiv 链接在原文稿中已存在换行拥挤；本题的契约是不改写其他内容，因此 Golden 保持该基线现象。若正式 Release 要求全文稿视觉达标，应先将此 Seed 改为“页序+已知版式修复”的新任务，不应在当前 Judge 中静默修改。
- LibreOffice/PDF 渲染能检查页序、页码、文本、图片和布局，但不能穷尽 PowerPoint 对页码域和占位符继承的全部行为；正式 Release 前建议在 PowerPoint 中打开并切换全部 6 页一次。
- 当前 L2 / Medium 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case 以及任务专属 `build_golden.py` / `solve.py`，核对六页标题、议程段落、页码对象与资产保持要求。
- 冻结输入、本地 Golden 与上游 Ground Truth 均已用 LibreOffice 完整渲染、制作六页联系图并逐页按原始分辨率查看。输入的议程、内容、总结和问答被打乱且 Engaging the audience 缺页码；Golden 按演讲逻辑重排并补齐 1–6。上游答案视觉顺序可用，但漏掉页码 3 且无关改写媒体与关系。
- 本题要求精确的六页顺序、议程顺序和模板页码修复，因此使用 `reference`。`V001` 查六页叙事流程，`V002` 查 Agenda 清单和分栏，`V003` 查右侧连续页码，`V004` 查每页主图/标题/正文/装饰保持，`V005` 查全稿主题一致和交付卫生。原生页码对象、关系 ID、备注、媒体哈希和包骨架继续由确定性 Judge 负责。
- Golden 可见风险集中在 Conclusion 页的三条彩色长链接，它们来自冻结输入且在 LibreOffice 中换行较紧；VLM 应判断候选是否保持，而不把源文稿已有的链接排版当作本次重排缺陷。
- Golden 的确定性 Judge reward 为 1.0；使用冻结的 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 独立运行 5 次，视觉分为 `[1.0, 1.0, 1.0, 1.0, 1.0]`，均值 1.0，总体标准差 0.0，满足重复运行标准差小于 0.05 的稳定性要求。首轮负例的议程和第 3 页页码偏移在实际渲染中过轻，Terra 仍给出 1.0，因此拒绝该负例。合格 v3 负例保留 Agenda -> Why Visuals Matter -> Engaging the Audience -> Effective Delivery Techniques -> Conclusion -> Q&A 六页顺序、所有影像和正文，但将 Agenda 五项整体右下移并改为暗红色，同时把第 3 页页码改为青色带外的大号横排红字；文件仍可打开且保持原生可编辑。其代码分为 0.0，Terra 视觉分为 0.678000，按 70/30 权重得到综合分 0.203400；视觉分严格低于 Golden，且高于空白或损坏文件。本题 7 次 Terra 调用均成功，虽有两次长尾等待，但没有基础设施失败；后续评分调用失败仍只记为基础设施失败，不按 0 分计入。
