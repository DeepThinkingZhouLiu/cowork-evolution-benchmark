# PPTArena Case 31 逐题适配记录

## 适配契约

- 角色：演示文稿版式编辑。
- 输入：1 页 16:9 AI 结论演示文稿。
- 真实意图：保持全部原文和格式，仅修改超出右边界的 `TextBox 1` 宽高，使全部文字在页内可见。
- 主交付物：`FixTextOverFlow_GroundTurthA.pptx`；沿用上游文件名中的 `GroundTurth` 拼写，避免改变产物契约。
- 不可妥协约束：左上角不变；813 个字符、文本格式和自适应设置不变；文本框完全在页内且有合理边距；标题、背景、主题和备注不改。
- 候选设计目标：L2 / Medium；正式难度字段保持 `unassigned`，等待真实 Agent Rollout 后定标。

## 上游来源与许可

- 来源：PPTArena Case 31: Fix Text Overflow。
- GitHub：<https://github.com/michaelofengenden/PPTArena.git>，commit `6f26a63ec1014b98c44990c7b58133b832070be2`。
- Hugging Face dataset commit：`20fb8902c74df2e97a2c31d288e4fa7f23405157`。
- 许可：GitHub README 与 Hugging Face 数据卡均标记 MIT；副本保存在本目录。
- 上游 Query、style target、输入和人工 Ground Truth 路径已保存在 `upstream-case.json`。
- 已完整阅读上游 `src/llm/judge.py` 和 `agent_bench/judge_predictions.py`，理解 smart diff、视觉判断和 Agent 产物定位链路；未运行依赖外部模型的网络 Judge。
- 本题沿用上游真实公开 PPTX 和人工任务意图，不是本项目自行编造的 Mock 输入。

## 已读入上下文的文件

| 文件 | 字节数 | SHA256 | 实际检查 |
|------|-------:|--------|----------|
| `data/input_files/FixTextOverFlow_TestA.pptx` | 8132227 | `183e662d0738577cee4c177dfcefc8620c664538e239b818aab47ee319b9a46e` | 1 页原尺寸渲染；文字、对象、几何、样式、备注、master/layout 和 8 个媒体审计 |
| 上游 `GroundTruth/FixTextOverFlow_GroundTurthA.pptx` | 8132127 | `2a034960c369db2d6db690be43d5aab35a9fadda169143552e5ddedd663d7890` | 1 页原尺寸渲染；几何、文本、格式、PDF 可见文本和包差异审计 |
| `solution/Golden/FixTextOverFlow_GroundTurthA.pptx` | 8116695 | `196da711fbaee730d8de7c2113af635ffe823e38867c3934518f338ccf82f378` | 1 页 150/300 DPI 渲染、原尺寸观察、PDF 文本提取和编辑性审计 |
| `adaptation/upstream-case.json` | 1102 | `3abb77f074b3959cd4b7c7ea8c500ee6c2ba0e262865a6b0d7765e96ba067e45` | Query、style target、路径和改造边界逐字段比对 |
| `adaptation/upstream-github-readme.md` | 4576 | `edb3b3b0cabe717f3880325ad4a9e991ebd8a64399d22b8fa3b76a813349be58` | 数据结构、规模、评测方式和许可标记 |
| `adaptation/upstream-hf-dataset-card.md` | 5021 | `163ac313d60d4f716b60313f5662584f522dcfae82de79bc3604a30abcaed4fa` | MIT 字段、公开状态和人工 Ground Truth 说明 |
| `adaptation/LICENSE-PPTArena-MIT.txt` | 1068 | `dbb2e3841e7e7e5901a005c4cc7480314bf4d1846dccf35d4574c37d9cb089b8` | 上游 MIT 许可文本 |
| 上游 `src/evaluation_pairs_refined.json` | 165586 | `17a584d1581a2464e4af8f06e1631dadc89c93a23624b54741e007ec6795850f` | Case 31 原始记录和等价答案要求 |
| 上游 `src/llm/judge.py` | 40829 | `82fa63bcc3206b6c83f34d15d5e49921b9413fa818052d474fa8174d74a3c6a2` | smart diff、Ground Truth 和视觉评分链路 |
| 上游 `agent_bench/judge_predictions.py` | 16100 | `c3e939afc5d14fdd2e2e0da43063314f1c94f31ac134e77e133e38b9ec1585308` | Agent 产物定位、多次采样和汇总入口 |

## 文稿与 Golden 审计

- 页面为 13.333×7.5 英寸，只有标题 `Conclusion` 和主文本框两个页面对象。
- 冻结输入的 `TextBox 1` 位于 x=0.6722、y=3.5694，宽 14.9667、高 2.5244 英寸，右边越出画布 2.3056 英寸。
- 本地 Golden 保持 x/y 不变，将宽高改为 10.653×3.433 英寸；右边距 2.008 英寸，底边距 0.498 英寸。
- 813 个字符的主文本、文本框边距、换行、`SHAPE_TO_FIT_TEXT`、段落和字体样式签名全部保持。
- LibreOffice 导出的 PDF 可完整提取 `Conclusion` 和全部主文本，原始越界文稿的 PDF 则会在多行右端截断单词。
- 上游人工 Ground Truth 几何与本地 Golden 接近，作为等价答案在独立 LibreOffice profile 中重放并获得 1.0。本地重建的原因是避免直接复制上游文件中与任务无关的 50 个 slide layout、slide master、custom XML 和 view properties 重写。
- `visual_qa.py` 对输入和 Golden 都只报告同一条上游既有 `underfilled_textbox` 通用警告：标题占其占位符高度约 17%。本题不擅自改标题占位符。

## Judge 设计

本地 Judge 不读取 Golden，不比较 Golden 哈希；它只使用冻结输入和任务契约。

- 验证 1 页、页面尺寸、ZIP 完整性和无宏。
- 根据原文识别主文本框，比对全部文字、段落、run 样式、边距、换行和自适应签名。
- 左上角容差为 0.02 英寸；文本框必须全部在页内，宽度至少 9.5 英寸、高度至少 3.0 英寸，并且相比输入明确减宽增高。
- Judge 用独立 LibreOffice profile 导出 PDF，再用 `pdftotext -layout` 验证标题和 813 字符主文本完整可见，因此仅在 OOXML 中保留但实际裁切的答案不会过关。
- 渲染后忽略主文本可编辑区，检查标题和背景区与冻结输入的像素相似度不低于 0.9985。
- 页面必须恰有两个原生文字对象；拒绝整页截图、额外对象、标题篡改、背景篡改和原媒体丢失。
- 独立合法变体使用 11.30×3.20 英寸，不复制 Golden 宽高，但仍能完整显示全部文本。

## 实际验证结果

| 样例 | 实际 reward | 结果 | 关键原因 |
|------|------------:|------|----------|
| 本地 Golden | 1.000000 | 通过 | 11 条判据全部通过 |
| 独立 11.30×3.20 英寸合法答案 | 1.000000 | 通过 | 文本完整、格式不变、位置不变且边距合理 |
| 上游人工 Ground Truth | 1.000000 | 通过 | 作为等价几何答案被正确接受 |
| fresh replay | 1.000000 | 通过 | 空输出目录回放产物 SHA 与已审计 Golden 一致 |
| 空交付 / 损坏 PPTX / 未修改输入 | 0.000000 | 拒绝 | 缺失、无法打开或仍然越界 |
| 移动左上角 / 仍超出右边 / 高度不足 | 0.000000 | 拒绝 | 位置、页内边界或可见文本门槛失败 |
| 改写文本 / 缩小字号 / 关闭换行 | 0.000000 | 拒绝 | 原文或文本格式签名失败 |
| 改标题 / 改背景 / 新增对象 / 新增页 | 0.000000 | 拒绝 | 非目标内容、结构或渲染保持失败 |
| 整页截图压平 / 夹带 `solution/Golden` | 0.000000 | 拒绝 | 可编辑性、结构或输出卫生失败 |
| 右边距仅 0.11 英寸 | 0.920000 | 中等 | 文本、格式、位置、可见性和可编辑性正确，仅合理边距分项失败 |

同一 Golden 独立判分 5 次的 reward 为 `[1, 1, 1, 1, 1]`，均值 1.0，总体标准差 0.0。

官方 Harbor 0.21 Oracle Docker 回放结果为 1.0，job 为 `pptarena031-oracle-fix-text-overflow-20260829`，trial 为 `pptarena-031-fix-text-overflow__yGAE52v`，0 个异常；容器产物 SHA256 为 `196da711fbaee730d8de7c2113af635ffe823e38867c3934518f338ccf82f378`，与已审计 Golden 一致。

主要命令：

    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task tasks/ppt/pptarena-031-fix-text-overflow
    PYTHONPATH=src python -m cowork_evo_bench.cli validate-task-assets tasks/ppt/pptarena-031-fix-text-overflow
    harbor run --path tasks/ppt/pptarena-031-fix-text-overflow --agent oracle --env docker --n-concurrent 1 --n-attempts 1 --force-build --yes --job-name pptarena031-oracle-fix-text-overflow-20260829

## 剩余风险

- 当前已通过 OOXML、python-pptx 语义对象、LibreOffice PDF/像素渲染和人工对照；正式 Release 前仍建议用 Microsoft PowerPoint 打开/保存抽检。
- 高 DPI 并发 Poppler/LibreOffice 渲染在本机偶发字形缓存干扰；Judge 使用每次独立 LibreOffice profile、150 DPI 单进程渲染和 PDF 文本提取避免该干扰。
- L2 / Medium 只是制题配比目标，正式难度必须经过冻结 Baseline Agent 重复 Rollout 后再写回。
