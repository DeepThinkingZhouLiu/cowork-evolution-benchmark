# PPTArena Case 82 逐题适配记录

## 适配契约

- 角色：竖版财会简历演示文稿编辑。
- 输入：4 页竖版 PPTX，共178 个原生页面对象和 41 个媒体部件。
- 真实意图：在每页内侧添加细、实线、中性深灰的统一页面边框。
- 主交付物：`AddSlideBorderGroundTruthB_AccountantCVResume.pptx`。
- 不可妥协约束：4/4 页都有唯一闭合可编辑边框；四边均匀内缩；边框样式跨页一致；原文字、图片、几何、裁剪、顺序和备注不变。
- 候选设计目标：L1 / Easy；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 82: Add Slide Border。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>。
- Git commit：`6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；许可和数据说明副本保存在本目录。
- 上游 Query、style target、输入和 Ground Truth 路径已逐字段保存在 `upstream-case.json`，未将它们伪称为本项目自编数据。

文稿含人像照片和外部简历模板素材。PPTArena 数据集的 MIT 标记不会自动消除第三方素材的权利风险，正式公开 Release 前必须单独复核授权或替换素材。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/AddSlideBorderTestB_AccountantCVResume.pptx` | 964243 | `bff930968572cce3f83f3bbb3eb26889090edfb7aceb681b109958ea78216ff2` | 4 页原尺寸渲染；178 个对象、41 个媒体、几何、文字、裁剪、顺序、备注和包部件审计 |
| `solution/Golden/AddSlideBorderGroundTruthB_AccountantCVResume.pptx` | 964752 | `b3cd1e8595deb342a8b2a7d2d3de323767d4754bb8f4037ac355334660be019b` | 上游人工 Ground Truth；4 页逐页原尺寸目视、OOXML 差异和媒体哈希审计 |
| `adaptation/upstream-case.json` | 930 | `98eb008564a80e45408eceaf382c9d3455b6429e0b70e6f02af96c26dcb851ba` | Query、style target、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | 公开状态、MIT 字段和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可原文 |

## 输入和上游 Ground Truth 审计

冻结输入为 4 页、8.2639×11.6944 英寸的竖版文稿，每页原对象数为 `[58, 57, 56, 7]`，共 41 个媒体和 81 个 ZIP 包成员。我逐页查看了原文稿与上游 Ground Truth 的 8 张原尺寸渲染页；原文稿自身在前两页存在局部文字重叠，Ground Truth 中完全保持，它不属于本题修改范围。

上游 Ground Truth 只改动 4 个 slide XML：每页末尾各增加一个原生无填充矩形，四边均内缩 0.2 英寸，线宽 0.72pt，颜色 `#555555`。其余 77 个包部件、41 个媒体、178 个原对象的文字、几何、裁剪、样式和顺序均保持，因此它符合本题契约，直接冻结为本地 Golden。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只读冻结输入，根据任务契约独立重算。

- 接受每页一个无填充矩形，也接受每页四条原生直线构成的闭合矩形，不锁死 Golden 实现。
- 四边内缩必须在 0.1–0.5 英寸，同页对边误差不超过 0.05 英寸，跨页位置差不超过 0.03 英寸。
- 线宽必须在 0.5–2pt，颜色必须是中性深灰且为实线；矩形必须无填充。
- 移除候选边框后，逐页比对原对象的内容、几何、旋转、裁剪、样式、顺序、版式和备注。
- 媒体哈希 Counter、包成员集和页面对象数独立检查，拒绝截图压平、替换图片、额外页和额外素材。
- LibreOffice 实际渲染冻结输入与候选文稿；遮罩边框线条区后，其余像素相似度必须不低于 0.9985。
- 在实际渲染图中重新检查四边深色线条的连续覆盖，并确认原对象没有越过边框。
- 移除合法边框后的内容 XML 和关键包部件字节保持只占 2 分，不会将合法答案锁死为某一个字节流。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 上游人工 Ground Truth / 本地 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立 0.3in、1.25pt、`#666666` 矩形答案 | 1.000000 | 通过 | 位置和样式不同于 Golden，但满足契约 |
| 独立 0.25in、1pt、`#444444` 四直线答案 | 1.000000 | 通过 | 使用 16 条原生直线，闭合、均匀且可编辑 |
| fresh replay | 1.000000 | 通过 | 从空输出目录重放，产物 SHA 与审计 Golden 一致 |
| 无害重打包 `[Content_Types].xml` | 0.980000 | 中等 | 功能、对象和渲染正确，只有 2 分字节级保持失败 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺文件、无法打开或没有边框 |
| 漏页 / 只改一页 / 某页重复边框 | 0.000000 | 拒绝 | 4/4 页唯一闭合边框失败 |
| 填充 / 过粗 / 过浅 / 红色 / 虚线 | 0.000000 | 拒绝 | 边框样式越界 |
| 贴页边 / 跨页不一致 / 左右不对称 / 小矩形 / 缺右线 | 0.000000 | 拒绝 | 边框位置、闭合或均匀性失败 |
| 改文字 / 移原对象 / 换图 / 答案标记 / 额外页 | 0.000000 | 拒绝 | 原内容保持、媒体、交付卫生或页数失败 |

包含 fresh replay 在内共 26 组正例、边界例和负例，全部与 `expectations.json` 一致。同一 fresh replay 产物独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena082-oracle-slide-borders-20260828`，trial 为 `pptarena-082-thin-dark-gray-slid__Eh4YCnX`，0 个异常；容器产物 SHA256 为 `b3cd1e8595deb342a8b2a7d2d3de323767d4754bb8f4037ac355334660be019b`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-082-thin-dark-gray-slide-borders
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-082-thin-dark-gray-slide-borders
    harbor run --path tasks/ppt/pptarena-082-thin-dark-gray-slide-borders --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes

## 视觉复核

- 已按 presentation-skill 流程渲染冻结输入、上游 Ground Truth 以及独立四直线合法变体。
- 已原尺寸查看输入和 Golden 的全部 4 页，以及四直线变体的首页和末页；边框闭合、线宽和颜色正常，未裁切、遮挡或移动原内容。
- Golden 的 4 页边框区外渲染相似度均为 0.999032，四边深色线条覆盖均为 1.0。
- 冻结输入的自动 QA 仅报告 1 条与任务无关的 `sparse_slide` 启发式警告；Golden 为 0 条警告。

## 剩余风险

- 正式公开 Release 前必须复核人像照片和简历模板的第三方权利；如无明确授权，需替换素材并保持同等边框编辑任务语义。
- LibreOffice/PDF 渲染可检查边框与布局，但不能穷尽 Microsoft PowerPoint 的全部编辑行为；正式 Release 前仍建议在 PowerPoint 中手动拖动一次边框，确认其可编辑性。
- 当前 L1 / Easy 只是制题配比目标，不是正式难度标签；必须经过冻结 Baseline Agent 的重复 Rollout 后才能写回 `difficulty_level` 和 `difficulty_tier`。
