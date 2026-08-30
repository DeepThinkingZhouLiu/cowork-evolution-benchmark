本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入、本地人工重建 Golden 和上游不合格 Ground Truth 现场生成。

- `valid-rgb-sampled`：从冻结输入独立制作，顶部 `FEFEFE`、底部 `B0C0D8` 均可在原图浅色像素中找到，不复制 Golden 的主题色参数。
- `valid-thin-accent-border`：在正确背景上为原图增加 1pt `accent1` 细边框，验证任务书明示允许的可选细边框。
- `repacked-content-types`：只给 `[Content_Types].xml` 增加无害空白，验证字节级保持只影响小额正向分。
- `upstream-unfit`：上游 PPTArena Ground Truth；其渐变结构正确，但 LibreOffice 中央 TIFF 图片被严重错渲染，用来验证 Judge 不会盲信上游答案。
- `solid-gray` / `background-rectangle`：验证必须是幻灯片原生线性渐变背景，不接受纯色或图形伪背景。
- `horizontal-gradient` / `pure-white-gradient` / `too-dark-bottom` / `yellow-gradient` / `reversed-gradient`：验证方向、亮度、色系、端点差异和上浅下深的确定性边界。
- `changed-title` / `moved-picture` / `changed-crop` / `missing-picture`：验证原文字、几何、裁剪和图片都必须保持。
- `thick-red-border`：验证可选边框不得过粗或偏离蓝色系。
- `hardcoded-marker` / `extra-slide`：验证答案路径泄漏、额外对象和页数变化不会过关。
