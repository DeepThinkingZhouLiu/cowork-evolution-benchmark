本题一对一改造自 OmegaUse-OfficeVal `officeval_051`。

- 原始 Query、中英文 Rubric、上游 Verifier 和 `Image 2.png` 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题在实际阅读原图后，逐区域重建 Multivariate Time Series、两类编码器、四段融合管线、异常评分、决策、训练目标和图例。
- 四色时序折线、5×3 特征矩阵、注意力节点、前馈网络、归一化曲线、异常分布曲线与全部连线均是原生可编辑对象。
- 交付页面共含 166 个原生对象，不含任何位图。
- Harbor Judge 不与 Golden 文件做字节比对，而是独立重算文本、矩阵尺寸、节点数、连线数、图片数、页面边界和真实渲染。

Golden 已经 LibreOffice 渲染为 1 页 PDF 并目视检查；本地 Judge 得分为 1.0。

## 视觉评分适配（visual-v1）

- 模型逐项读取了 `instruction.md`、`task.toml`、确定性 `rubric.json`、完整 `tests/judge.py`、输入清单、来源记录、中英文上游 Query/Rubric、本题改造说明和整份专用 `solve.py`；本题没有 `tests/expectations.json`，未虚构该文件。
- 唯一输入 `Image 2.png` 已按原始分辨率查看：它以青、紫、橙、绿、粉、蓝六色区分时序输入、双编码器、四阶段融合、异常评分、决策、训练目标和图例。Golden 已真实渲染为 38.4×20 cm 的单页图像，并目视检查全部 166 个原生对象的可见结果。
- 上游要求“一模一样”且指定按图重建，因此本题使用 `reference`；VLM 同时查看候选与 Golden，并以原图的区域层级、连接关系和模块配色为准。
- `V001` 检查八区域宏观布局；`V002` 检查时序输入、双 GRU、特征矩阵与分流；`V003` 检查四阶段橙色融合管线；`V004` 检查右侧异常评分与决策轴；`V005` 检查训练目标、辅助回路和图例；`V006` 检查整页科研图的配色、线条、字号与可读性。文本、对象数、位图禁用和页面尺寸继续由确定性 Judge 负责。
- Golden 的可见风险包括：左侧时序卡较原图更高且右上留白较多；融合区卡内图标简化，四阶段与原图的卡间节奏有小差异；训练区与右侧决策区的字号相对小。VLM 需按真实渲染评分，不得因其为 Golden 自动给满分。
- Golden 保留确定性得分 `1.0`；the VLM evaluator 连续 5 次视觉得分均为 `1.0000`，均值 `1.0000`，总体标准差 `0.0000`，满足 `std < 0.05`。
- 中等可见负例仅修改本题对象：将 Self-Attention 标题下移到节点图上，将 Thresholding 卡下移至 Normal / Anomaly 卡上，并将 Auxiliary Flow 图例左移到虚线上。文件仍可打开，八区域和主流程保留，但三处遮挡分别直接对应 `V003`/`V004`/`V005`。该负例视觉得分 `0.7456`，严格低于 Golden 均值且高于空白/损坏产物；70/30 合成得分为 `0.9237`。
- 校准证据保存在忽略目录 `.coworkevo/vlm-calibration/ppt/officeval-051-image-to-editable-anomaly-detection-flow/`，包含 5 次 Golden、候选/参考渲染图和中等负例结果。
