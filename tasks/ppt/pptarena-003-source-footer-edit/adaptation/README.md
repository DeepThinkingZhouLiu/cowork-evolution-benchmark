# PPTArena Case 3 逐题适配记录

## 适配契约

- 角色：演示文稿编辑。
- 输入：13 页 VerifyMe 2025 年第一季度投资者电话会演示文稿。
- 真实意图：在第 5 页页脚增加指定来源引用，同时保护其他内容。
- 主交付物：`VerifyMe_Q1_2025_with_source.pptx`。
- 不可妥协约束：引用文字精确、原生可编辑、位于第 5 页页脚，不改动原有 13 页内容。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout。

## 上游来源与许可

- 来源：PPTArena `Case 3: Add Footer Text Box`。
- GitHub：`https://github.com/michaelofengenden/PPTArena.git`。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在 `upstream-github-readme.md`、`upstream-hf-dataset-card.md` 和 `LICENSE-PPTArena-MIT.txt`。
- 上游 Query、style target、输入和 Ground Truth 映射完整保存在 `upstream-case.json`。
- 上游数据卡明确称 Ground Truth 为人工编辑目标；本题不伪称为自行编造的业务数据。

原文稿含 VerifyMe 品牌和公开投资者沟通内容。本候选题依据上游 MIT 数据集授权保留完整归因；在正式 Release 前仍需纳入项目级第三方素材法务复核。

## 已读入上下文的文件

| 文件                                      | 字节数 | SHA256                                                           | 实际检查                                       |
|-------------------------------------------|-------:|------------------------------------------------------------------|------------------------------------------------|
| `data/input_files/VerifyMe_Q1_2025.pptx`  | 471140 | `01238cae60124a40312629aa2f8ec5c6735ef4e38bc56f0f85691336cf1fe400` | 13 页渲染、联系表、第 5 页对象与文字展开检查       |
| `solution/Golden/VerifyMe_Q1_2025_with_source.pptx` | 448791 | `8acc4b32d9bc7582a3d2c560ec91c73b29d825371b0c8b21863a8a4837511fe7` | 13 页渲染、联系表、每页原尺寸查看、OOXML 目标形状检查 |
| `upstream-case.json`                      |   1261 | `c426e23d1e9e449a571971aba3f315cf87b0bb26a0268413bc005b23240616d1` | Query、style target、enhancement notes 逐字阅读    |
| `upstream-github-readme.md`               |   4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记                    |
| `upstream-hf-dataset-card.md`             |   5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开/非 gated 状态、人工 Ground Truth 说明  |

## 上游答案审计与适配决策

上游 Golden 在第 5 页新增了一个文本框：

- 文字：`Source: www.VerifyMe.com`。
- 位置与尺寸：`x=0.500, y=7.000, w=12.333, h=0.250` 英寸。
- 显式字号：10 pt。
- 对象类型：PowerPoint 原生可编辑文本框，不是图片。

13 页中，除第 5 页外的 12 页 LibreOffice 渲染文件与原输入逐字节一致；第 5 页分隔线上方的渲染也一致。上游 Golden 经 PowerPoint 保存时重写了大量无关 OOXML 关系和元数据，因此本题明确不做整包字节比较，避免把合法的 Office 重保存误判为错误。

本地 Judge 改为独立检查：指定文本、页码、原生可编辑性、页脚几何、原文保持、原图媒体哈希、非目标页渲染相似度、第 5 页主内容渲染相似度和 PDF 文本可提取性。Judge 不读取 Golden，也不比较 Golden 哈希。

## 实际验证结果

| 样例                         | 实际 reward | 结果 | 关键原因                                      |
|------------------------------|------------:|------|-----------------------------------------------|
| 上游人工 Golden              |    1.000000 | 通过 | 11 条判据全部通过                            |
| fresh replay                 |    1.000000 | 通过 | `solve.sh` 从空输出目录重放人工 Golden        |
| 空交付                       |    0.000000 | 拒绝 | 缺失指定交付物                                |
| 损坏 PPTX                    |    0.000000 | 拒绝 | ZIP/OOXML 无法打开                           |
| 未修改输入                    |    0.000000 | 拒绝 | 缺失指定引用                                  |
| 错误 URL                      |    0.000000 | 拒绝 | 引用不精确且夹带额外文本                        |
| 引用写在第 4 页                |    0.000000 | 拒绝 | 第 5 页缺失，且改动其他页                      |
| 将 `22.6%` 改为 `99.9%`    |    0.000000 | 拒绝 | 篡改原有关键数值                                |
| 第 5 页整页截图               |    0.000000 | 拒绝 | 原生文本缺失、媒体集变更、新增全页图              |
| 重复两次引用                  |    0.000000 | 拒绝 | 引用次数和原文保持失败                         |
| 夹带 `solution/Golden` 标记 |    0.000000 | 拒绝 | 新增了未授权的泄漏文本                          |
| 引用文字正确但位于页外          |    0.773585 | 中等 | 内容与原文保持正确，但位置、可读和渲染交付失败       |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 `1.0`，总体标准差 `0.0`。

官方 Harbor 0.21 Oracle 容器回放结果为 `1.0`：环境构建、`/solution` 执行、`/app/output` 产物收集、Judge 和 `/logs/verifier/reward.txt` 全部贯通。首轮直连 Debian 源因容器网络超时中止；将 Dockerfile 固定到清华 Debian/PyPI 镜像后，官方回放在 81 秒内无异常完成。

主要命令：

```bash
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-003-source-footer-edit
PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-003-source-footer-edit
harbor run --path tasks/ppt/pptarena-003-source-footer-edit --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes
```

## 视觉复核

- 已查看原输入和 Golden 的 13 页联系表。
- 已逐页查看 Golden 的 13/13 页原尺寸渲染图。
- 第 5 页新增引用完整可读，位于分隔线下方；未遮挡主内容和页码。
- 其余页面未出现新增占位符、空白页、裁切、离页或新的遮挡。
- 原文稿第 3、11 页文本密度高，但这是上游输入的既有特征，本题明确要求保持，不把其误当作候选答案的新缺陷。

## 剩余风险

- 正式公开 Release 前仍需对原文稿内的 VerifyMe 品牌与第三方投资者材料做项目级法务复核；当前依据是 PPTArena 公开数据卡和仓库的 MIT 标记。
- Linux 容器不含演示文稿原始的 DM Sans 字体，LibreOffice 会使用固定备用字体。源文件和候选文件在同一容器中渲染，相对保持判定稳定；正式 Release 可再冻结合法的 DM Sans 字体包以减少跨平台差异。

## 视觉评分适配与校准

制题模型已逐项读取 `instruction.md`、`task.toml`、`rubric.json`、完整 `tests/judge.py`、`tests/expectations.json`、上游 Case 记录、完整解题脚本、原始输入和人工 Golden；并渲染了输入与 Golden 的 13 页联系图，逐页查看后重点放大第 5 页页脚。本题是精确的局部编辑与保真任务，因此采用 `reference` 模式：VLM 对照候选和人工 Golden，判断 13 页原稿保真、第 5 页引用的安全位置、页脚层级及原生融合。

视觉 Rubric 共 4 条、按本题真实交付特征逐条编写，不替代原代码 Judge：代码 Judge 继续检查精确字符串、页码、原生可编辑对象、几何边界和原文/媒体保留；VLM 只补充检查联系图中真实可见的对齐、留白、层级、背景连续性与专业完成度。综合分仍按代码分 70% + VLM 分 30% 计算，两个分项单独保存。

固定 Codex CLI 0.148.0、ZCloud `gpt-5.6-terra` 对 Golden 的 5 次独立真实视觉评分均为 1.000，均值 1.000、总体标准差 0.000（低于 0.05）。首次仅轻微放大第 5 页来源文字的负例被 VLM 判为 1.000，未作为校准负例；返工负例保持 13 页和正确引用内容，把来源移到更靠近分隔线的位置并改为 15pt 红色粗体，仍处于页脚安全区、可打开且完整可读。该负例视觉分为 0.79518，严格低于 Golden 且高于空白/损坏交付；确定性分仍为 1.000，按 70%/30% 合成为 0.938554。一次 Golden 调用三次超时后按基础设施失败留痕，随后在新目录重试成功，未计 0。
