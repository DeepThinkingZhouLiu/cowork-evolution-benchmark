# Case 64 测试夹具

`tests/build_fixtures.py` 仅针对本题冻结输入和本地 Golden，构造 ZIP 压缩方式不同但 Office 语义完全相同的合法变体，以及配色顺序错误、非 Okabe–Ito 色、错数值、错系列名、水平条形图、删图例、重复图表、形状覆盖、改说明文字、移图表、整页压平、硬编码标记和额外输出等样例。

`golden-slide1.png` 只用于构造整页压平负例，不被 Judge 当作答案或视觉比对基准。
