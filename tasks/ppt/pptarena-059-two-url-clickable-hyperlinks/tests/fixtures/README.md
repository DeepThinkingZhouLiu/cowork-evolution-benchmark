本目录不手工存放大量 PPTX；由 `../build_fixtures.py` 基于冻结输入、本地人工重建 Golden 和上游不合格 Ground Truth 现场生成。

- `valid-explicit-blue`：从冻结输入独立制作，使用不同关系 ID，并为两个 URL 显式设置常见蓝色下划线，验证允许默认 hyperlink 样式。
- `valid-split-url-runs`：将每个 URL 分成两个 run，但同一 URL 的所有字符指向同一正确关系，验证 Judge 按字符语义判分而不锁死 run 数。
- `repacked-content-types`：只给 `[Content_Types].xml` 增加无害空白，验证字节级保持只影响小额正向分。
- `upstream-unfit`：上游 Ground Truth 已正确添加超链接，但同时移动整页背景大图导致左右裁切，验证 Judge 不会盲信上游答案。
- `style-only`：只把 URL 改为蓝色下划线，但没有功能 hyperlink。
- `only-berkeley` / `only-illinois`：验证两个 URL 都必须完整可点击。
- `swapped-targets` / `wrong-http-scheme` / `internal-target-mode`：验证链接目标和外部关系属性必须精确。
- `whole-sentence-linked` / `surrounding-linked` / `shape-level-link`：验证只能链接 URL 字符，不能链接整句、周围文字或整个对象。
- `wrong-font-size` / `changed-url-text` / `changed-title`：验证字号和所有原文字必须保持。
- `moved-background` / `changed-image`：验证整页背景大图几何与原媒体必须保持。
- `hardcoded-marker` / `extra-slide`：验证答案路径泄漏、额外对象和页数变化不会过关。
