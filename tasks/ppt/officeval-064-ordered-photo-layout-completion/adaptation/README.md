本题一对一改造自 OmegaUse-OfficeVal `officeval_064`。

- 原始 Query、中英文 Rubric、上游 Verifier、2 页 PPTX 模板和 12 张真实输入照片已原样保留；所有输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际阅读模板和 12 张图片后，按照 1 -> 12 的顺序逐框放置：第 1 页为上大下双小，第 2 页为左双小右大。
- 每张照片均为独立图片对象，使用原始图片字节和居中对称裁剪，不做拉伸或重编码。
- 两页原有 46 个对象的文本和几何位置保持不变；标题外围新增两个无填充的浅蓝色可编辑矩形，形成克制的双线边框。

Golden 已经 LibreOffice 渲染为 2 页 PDF 并逐页目视检查；上游 Verifier 得分 23/23，本地独立 Judge 得分为 1.0。
