本题一对一改造自 OmegaUse-OfficeVal `officeval_043`。

- 原始中文 Query、英文 Query、中英文 Rubric 与上游 Verifier 已原样保留在本目录和 `tests/`。
- 原始 PNG 逐字节保留，SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题 Golden 由模型实际读取名片图后逐页、逐区域重建：两页文字、外框、分栏、标识、联系图标、业务图标与底部色带均为原生可编辑对象。
- 只有二维码按原 Query 保留为来源图裁剪小图；页面不嵌入整张参考图。
- 上游 Verifier 依赖大量非语义形状启发式，对视觉上等价的原生组合形状会漏判。Harbor Judge 因此按可确定的页数、文本、几何、图片占比、原生对象数和真实渲染结果独立复算。

Golden 已经 LibreOffice 渲染为 2 页 PDF 并逐页目视检查；本地 Judge 得分为 1.0。
