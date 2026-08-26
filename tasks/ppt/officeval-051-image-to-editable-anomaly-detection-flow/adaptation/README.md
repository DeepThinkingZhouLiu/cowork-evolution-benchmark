本题一对一改造自 OmegaUse-OfficeVal `officeval_051`。

- 原始 Query、中英文 Rubric、上游 Verifier 和 `Image 2.png` 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题在实际阅读原图后，逐区域重建 Multivariate Time Series、两类编码器、四段融合管线、异常评分、决策、训练目标和图例。
- 四色时序折线、5×3 特征矩阵、注意力节点、前馈网络、归一化曲线、异常分布曲线与全部连线均是原生可编辑对象。
- 交付页面共含 166 个原生对象，不含任何位图。
- Harbor Judge 不与 Golden 文件做字节比对，而是独立重算文本、矩阵尺寸、节点数、连线数、图片数、页面边界和真实渲染。

Golden 已经 LibreOffice 渲染为 1 页 PDF 并目视检查；本地 Judge 得分为 1.0。
