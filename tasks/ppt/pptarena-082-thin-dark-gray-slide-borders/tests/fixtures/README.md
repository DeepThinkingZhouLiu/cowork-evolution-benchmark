本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入和已审计的上游人工 Golden 现场生成。

- `valid-gray-rectangle`：从冻结输入独立制作，每页使用 0.3in 内缩、1.25pt、`#666666` 的无填充矩形，不复制 Golden 的精确参数。
- `valid-four-lines`：每页用 4 条原生直线构成 0.25in 内缩、1pt、`#444444` 的闭合边框，验证 Judge 不锁死为单一矩形实现。
- `repacked-content-types`：只给 `[Content_Types].xml` 增加无害空白，验证字节级保持只影响小额正向分。
- `missing-slide4` / `only-slide1` / `double-slide2` / `missing-right-line`：验证 4/4 页都必须恰好有一个闭合边框。
- `filled-rectangle` / `too-thick` / `too-light` / `red-border` / `dashed-border`：验证无填充、线宽、深灰颜色和实线要求。
- `edge-touching` / `nonuniform-pages` / `lopsided-margin` / `small-content-box`：验证内缩范围、跨页一致、对边均匀和全页框定。
- `changed-text` / `moved-original` / `changed-image`：验证原文字、几何和 41 个媒体必须保持。
- `hardcoded-marker` / `extra-slide`：验证答案路径泄漏、额外对象和页数变化不会过关。
