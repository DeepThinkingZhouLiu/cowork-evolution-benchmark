本题一对一改造自 OmegaUse-OfficeVal `officeval_059`。

- 原始 Query、中英文 Rubric、上游 Verifier、4 页 Word 申报书和 18 页 PPT 已原样保留；输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际阅读 Word 全部 4 页、定位最后一行右侧单元格，并检查 PPT 首页及 18 页结构后，加入了真实 OOXML 外部超链接。
- 链接目标严格为同目录相对文件名 `青梧谷灵MR共创计划.pptx`；随附 PPT 与输入逐字节一致。
- Word 只有 `word/document.xml` 和 `word/_rels/document.xml.rels` 两个包成员发生变化，原正文、表格、页眉、页码和边框均保留。
- 上游 Rubric/Verifier 错把真实项目“青梧谷灵”写成“海潮艺境”，并要求不存在的 `海潮艺境MR共创计划.pptx`。本地 Judge 以真实输入文件、原 Query 和可点击相对链接为真值，因此没有迎合这处错误。

Golden 已经 LibreOffice 渲染为 4 页 PDF 并逐页目视检查；本地独立 Judge 得分为 1.0。上游 Verifier 因上述名称冲突只能得到 4/19，此差异已如实记录。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分逐字节保存在 `rubric.v1.json` 和 `tests/judge_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 不再因为单个链接文字、样式、关系属性或 PPT 页面错误把整题硬归零。Judge 分别核验 18 页 PPT 内容与媒体、目标行的唯一超链接、关系类型与安全相对路径、小五号蓝色下划线样式、原 Word 包成员、四页渲染、25 行申报表和无危险外链，并按真实原子检查连续给分。
- v2 Golden 与两次 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。30 个 v2 负例包含 8 个核心失败（均为 0）和 22 个有效局部错误（`0.148512–0.995172`）；“文字仍在但真实链接被移除”的中度负例为 `0.649803`，只保留极少有效内容的严重负例为 `0.148512`。
- 随附的 18 页 PPT 使用 reference-mode 五项视觉 Rubric，逐页检查顺序与骨架、封面项目身份、竞争/技术/运营高密度页、跨页卡片与图片一致性，以及第 18 页收束版式。Codex CLI `0.148.0` + `your-vision-model` 的五次 Golden 视觉分均为 `1.0`，总体标准差为 `0`。
- 任务专属视觉中等负例只在第 1 页覆盖项目标题和副标题，仍可打开、仍有 18 页且其余 17 页不变；确定性分为 `0.980690`、视觉分为 `0.650000`，70/30 组合分为 `0.881483`，低于 Golden 且高于无效产物。
- 本轮没有 VLM 基础设施失败；仓库中也未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，因此没有伪造 Harness 结果。
