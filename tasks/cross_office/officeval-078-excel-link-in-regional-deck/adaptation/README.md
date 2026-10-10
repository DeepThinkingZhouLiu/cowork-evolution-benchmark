本题一对一改造自 OmegaUse-OfficeVal `officeval_078`。

- 原始 Query、中英文 Rubric、上游 Verifier、10 页 PPT 和真实 CRM 工作簿已原样保留；输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际检查第 8 页全部 14 个原对象及右侧空白区域，并读取工作簿的 3 张工作表和记录维度后，设计了可编辑的 Excel 绿色圆角按钮。
- 按钮位于 x=27.0 cm、y=7.4 cm，尺寸 4.6 × 2.4 cm；不遮挡表格、标题、页脚和右下圆形。
- 超链接绑定在整个按钮形状上，目标严格为同目录相对文件名 `CRM_pipeline_restyled_reordered.xlsx`；随附工作簿与输入逐字节一致。
- 原 PPT 只有 `ppt/slides/slide8.xml` 和 `ppt/slides/_rels/slide8.xml.rels` 两个包成员变化，其他页面和原对象保持不变。

Golden 已经 LibreOffice 渲染为 10 页 PDF 并逐页检查；上游 Verifier 得分 10/10，本地独立 Judge 得分为 1.0。

## v2 连续与视觉评分改造（2026-09-04）

- 原严格评分逐字节保存在 `rubric.v1.json` 和 `tests/judge_v1.py`；`SCORING_VERSION=v1|v2` 可显式选择，默认使用 v2。
- v2 逐项核验三张 CRM 工作表及 857 条数据记录、按钮形状级点击区与两级文字、九项几何边界、安全相对关系、十页原对象与包成员、第 8 页渲染、绿色圆角样式和无额外交付。只有核心文件失败或原样交回输入 PPT 时硬归零，局部错误按本题原子单元连续给分。
- v2 Golden 与两次 fresh replay Reward 均为 `1.0`，v1 Golden 仍为 `1.0`。33 个 v2 负例包含 8 个核心失败（均为 0）和 25 个有效局部错误（`0.146621–0.998082`）；“按钮可见但关系断裂”的中度负例为 `0.793810`，仅剩一页标题的严重负例为 `0.146621`。
- 十页 PPT 使用 reference-mode 五项视觉 Rubric，分别检查跨页主题保留、第 8 页七行表格与页脚、按钮相对表格/圆形的留白、Excel 绿色圆角与两级白字，以及可直接汇报的专业度。Codex CLI `0.148.0` + `your-vision-model` 的五次 Golden 视觉分均为 `1.0`，总体标准差为 `0`。
- 任务专属视觉中等负例把第 8 页按钮左移到表格上方并改为红色，文件仍可打开且其余页面不变；确定性分为 `0.962083`、视觉分为 `0.625000`，70/30 组合分为 `0.860958`，低于 Golden 且高于无效产物。
- 本轮没有 VLM 基础设施失败；仓库中未发现可归属本题的 Claude Code、OpenCode 或 OpenHands 候选产物，因此没有伪造 Harness 结果。
