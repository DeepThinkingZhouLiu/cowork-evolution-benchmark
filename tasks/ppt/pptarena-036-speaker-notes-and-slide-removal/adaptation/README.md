# PPTArena Case 36 逐题适配记录

## 适配契约

- 角色：演示文稿结构编辑。
- 输入：1 份 10 页、含 8 张教学图片的漫画转场课件。
- 真实意图：识别原第 2 页的 8 个长文本框，将其逐一转移到后续主题页的讲者备注，再删除临时文本页。
- 主交付物：`AddSpeakerNotes_GroundTruth.pptx`。
- 不可妥协约束：9 页；8 段原文完整且唯一映射；封面无备注；保留页的内容、顺序、对象、几何、文本样式和图片不变。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 36: Add Speaker Notes。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>，commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在本目录。
- 上游 Query、style target、enhancement note、输入和人工 Ground Truth 路径已原样保存在 `upstream-case.json`。
- 已完整阅读上游 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，理解 smart diff、视觉判断、产物定位和重复采样链路；未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开 PPTX、8 段原始备注文字和人工编辑意图，不是本项目自行编造的 Mock 输入。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/AddSpeakerNotes_TestA.pptx` | 948105 | `a81cd08f0bdd05079ba350c1fa32b95db8111e504ae657613707e69abf8fd201` | 10 页渲染；逐页标题、对象、备注、页面关系、布局和 8 个媒体审计 |
| 上游 `GroundTruth/AddSpeakerNotes_GroundTruth.pptx` | 945871 | `7f9e401dd23cdd25665f9202caa42488601d65598be90d9e30f59017a8ea4da4` | 9 页渲染；8 段备注、页顺序、对象签名和包差异审计；本地 Judge 得分 1.0 |
| `solution/Golden/AddSpeakerNotes_GroundTruth.pptx` | 943293 | `314fe0f387211ee4874d0be6bb2f6d90d48611e8fe841ee9354d46fea743d03d` | 任务专属 OOXML 重建；9 页、9 个备注页、9 个幻灯片关系和应用元数据审计 |
| `adaptation/upstream-case.json` | 1879 | `303cc1a486039e6af794607fe23efa7d5ce8d8422babfd7e6f68a9bc26226c4a` | Query、style target、enhancement note、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和 MIT 标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| 上游 `src/evaluation_pairs_refined.json` | 165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 36 原始记录和等价答案要求 |
| 上游 `src/llm/judge.py` | 40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分链路 |
| 上游 `agent_bench/judge_predictions.py` | 16100 | `3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样和汇总入口 |
| `tests/fixtures/golden-slide3.png` | 559202 | `c40cbeee0c0d1cc9ffea585cbf23b828e6910a3c8f1853981910f777c6a5b1d8` | 仅用于构造整页截图压平负例 |

## 原题文本与页面映射

| 原第 2 页对象 | 原始页 | 最终页 | 目标主题 | 规范化字符数 |
|-------------------|---------:|---------:|----------|---------------:|
| shape id 4  | 3  | 2 | `Panel to Panel Transitions` | 270 |
| shape id 6  | 4  | 3 | `Moment To Moment Transitions` | 607 |
| shape id 8  | 5  | 4 | `Action to Action Transitions` | 274 |
| shape id 10 | 6  | 5 | `Subject to Subject Transitions` | 1080 |
| shape id 12 | 7  | 6 | `Scene to scene Transitions` | 948 |
| shape id 14 | 8  | 7 | `Aspect to Aspect Transitions` | 1033 |
| shape id 16 | 9  | 8 | `Non Sequitur Transitions` | 1066 |
| shape id 18 | 10 | 9 | `Group Activity: Transitions and empathy / Identification` | 356 |

上表中的所有文本都是从冻结输入的原对象读取；没有补写、润色或纠正上游原文中的拼写和语法。

## Golden 构建与审计

- 冻结输入为 10×5.625 英寸，10 页，包含 8 个图片媒体和 10 个空备注页。
- `build_golden.py` 不是通用生成器。它只处理本题锁定哈希的文件，按已审计的 8 对映射复制原始段落 XML，删除原第 2 页、其空 notesSlide 和关系，并同步 `Slides`、`Notes`、`Slide Titles` 计数。
- 保留页对应的 slide XML、slide rels、layout、master、嵌入字体和媒体原样复用；`python-pptx` 对 9 个保留页重算的页面对象、几何、文字、段落、run 样式和图片哈希签名与原稿全部一致。
- 上游人工 Ground Truth 的备注映射也正确，并在本地 Judge 获得 1.0。本地 Golden 仍从冻结输入重建，目的是减少非目标 XML 重写，不是因为上游答案映射错误。
- 封面备注为空；后续 8 页备注规范化后与原文分别为 270、607、274、1080、948、1033、1066 和 356 个字符。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只使用冻结输入和本题契约。

- 校验 ZIP 完整性、无宏、关系目标存在、9 个幻灯片关系和 9 个 notesSlide。
- 从冻结输入原第 2 页动态重算 8 段真值；每段必须完整出现在唯一的正确目标页备注中。
- 要求 9 个标题按已审计顺序出现；封面备注为空，临时页长文不得作为可见页面内容残留。
- 对原第 1、3–10 页与候选第 1–9 页逐页比较对象数、z-order、几何、类型、占位符、文本、段落、run 字体样式和图片哈希。
- 要求媒体哈希多重集合与输入完全一致，且不得出现覆盖全页的图片。
- 用每次独立 LibreOffice profile 转 PDF，校验 9 页、9 张非空渲染图和 9 个页面标题可提取。
- 允许只改变空白和换行的等价备注；多加演讲断言会丢失 6 分内容纯净度，但不触发核心映射门槛。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 只改变一处备注换行的合法答案 | 1.000000 | 通过 | 规范化文本、页顺序和保留页签名全部正确 |
| 上游人工 Ground Truth | 1.000000 | 通过 | 作为等价包结构和备注样式被正确接受 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物 SHA 与已审计 Golden 一致 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或仍为 10 页 |
| 删错页 / 调换保留页顺序 / 额外加页 | 0.000000 | 拒绝 | 页数、标题顺序或对象映射门槛失败 |
| 缺备注 / 串页 / 截断 / 重复 / 封面误加备注 | 0.000000 | 拒绝 | 8 段文本的完整性、唯一性或目标页失败 |
| 改页面文字 / 移动对象 / 新增对象 / 删除图片 | 0.000000 | 拒绝 | 保留页签名或媒体多重集合失败 |
| 将一页压平为整页截图 | 0.000000 | 拒绝 | 对象结构、媒体和可编辑性失败 |
| 夹带 `solution/Golden` | 0.700000 | 不通过 | 输出卫生失分且触发严重错误扣分 |
| 在一页备注多加一句演讲提示 | 0.940000 | 中等 | 原文和映射完整，仅备注纯净度失分 |

同一最终 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

最终官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena036-oracle-speaker-notes-final-20260829`，trial 为 `pptarena-036-speaker-notes-and-s__QNG8gN4`，0 个异常；容器产物 SHA256 为 `314fe0f387211ee4874d0be6bb2f6d90d48611e8fe841ee9354d46fea743d03d`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-036-speaker-notes-and-slide-removal
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-036-speaker-notes-and-slide-removal
    harbor run --path tasks/ppt/pptarena-036-speaker-notes-and-slide-removal --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --job-name pptarena036-oracle-speaker-notes-final-20260829

## 视觉检查与剩余风险

- 输入 10 页和 Golden 9 页均已全页渲染并人工检查；页面和图片完整，无空白页或丢图。
- `visual_qa.py` 在原稿第 5–7 页和删页后第 4–6 页报告相同的 3 条 `underfilled_textbox` 通用警告；这些是上游占位符高度，本题不擅自修改。
- 这份由 Google Slides / PowerPoint 产生的文稿包含嵌入字体与继承文本样式。LibreOffice 在幻灯片数量从 10 变为 9 后，会对部分继承文本产生与原 10 页文稿不同的排版；上游 Ground Truth 和本地 Golden 都呈现同样的 LibreOffice 差异，而保留页的 OOXML 对象、几何、字体、字号和媒体签名全部未变。
- 因此本地 Judge 以保留页的语义/样式签名为保持真值，并要求 LibreOffice 可渲染、无空页、标题可提取，但不强行要求删页前后的 LibreOffice 像素一致。正式 Release 前应在 Microsoft PowerPoint 中打开/保存抽检；如果最终 Baseline 只允许 LibreOffice，则应将该兼容风险纳入题目发布决策。
- L2 / Medium 只是制题配比目标，正式难度必须经过冻结 Baseline Agent 重复 Rollout 后再写回。

## 视觉评分适配（visual-v1）

- 制题模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、本改造记录、上游 Case、任务专属 `build_golden.py` / `solve.py`、冻结输入和本地 Golden，核对了原第 2 页的 8 段真实文字、9 个目标页标题、备注映射、页面对象签名和 8 个媒体。
- 冻结输入 10 页和 Golden 9 页均已全部渲染、制作联系图并逐页亲眼查看：输入第 2 页是八块长文堆叠的临时备注页，Golden 删除它后保留了封面、六种分镜转场概念/案例和 Group Activity 页的蓝色标题、黑色正文、漫画图片与顺序。
- 本题是只允许备注迁移和删除指定页的结构编辑，因此使用 `reference`；参考用于判断保留页的标题、正文、图片取景、几何和留白是否漂移。专属视觉 Rubric 分别检查 9 页教学序列与临时页消失、保留页构图、漫画清晰度/裁切、教学文本可读性和整稿风格干净度。
- 讲者备注是不在普通幻灯片渲染中可见的包内状态，8 段原文的完整性/唯一性/主题映射、封面空备注、9 个 notesSlide 关系、保留页对象/样式/几何、媒体哈希和原生可编辑性继续由确定性 Judge 负责；VLM 不根据不可见备注猜测分数，只补充联系图与单页中的可见顺序、裁切、重叠、留白和专业完成度。综合分按 70% 代码分 + 30% 视觉分计算，两个分项单独保存。
- Golden 使用固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 真实独立评分 5 次，视觉分为 `[1.000, 1.000, 1.000, 1.000, 1.000]`，均值 `1.0000`，总体标准差 `0.0000`，小于 `0.05`；5 次综合分均为 `1.0000`。
- 中等可见负例从 Golden 手工定向改造：保留 9 页、八段讲者备注、所有原文、标题、图片、媒体和原生对象，只将第 9 页 Group Activity 漫画放大并左移，让图片明显遮住问题文本的右半部；已全页渲染并亲眼确认文稿仍可打开，前 8 页完好，最后一页仍能辨认标题、问题列表和漫画，但图文边界与可读性有清楚的中等缺陷。因确定性 Judge 对非目标页几何采用严格保持门槛，其代码分为 `0.0000`；视觉分 `0.7100`、综合分 `0.2130`，严格低于 Golden 均值且高于空白/损坏交付。
- 本轮部分调用触发了 Provider 长尾和配置内自动重试，但最终全部成功并写入完整结果；无未解决基础设施失败，未将任何失败冒充为 0 分。
