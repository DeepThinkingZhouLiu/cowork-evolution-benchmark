# PPTArena Case 93 逐题适配记录

本题逐页读取了 PPTArena Case 93 的 Query、`style_target`、原 6 页 PPTX 和上游 Ground Truth，没有用通用生成器批量编题。

上游答案的深蓝标题、深灰正文、7 页递增进度条、第 4 页原生簇状柱形图及 14.38/14.98/15.10 三个数值、第 5 页两张原图的重排均正确。但它的 Bibliography 除 Copernicus 外还显示 UN、Wikipedia、NASA、EPA、Climate.gov 和 FAO 网址，这 6 条并未出现在本题上游数据记录中，且显示网址与实际超链接目标不一致。本地 Golden 因此逐项删除无来源条目，只保留 `upstream-case.json` 明确给定的完整 Copernicus URL，并使显示文字、下划线和点击目标一致。

Judge 不比较 Golden 哈希。它独立读取图表缓存中的类别和数值、两张原图的二进制哈希、图注空间关联、每页进度比例、主题文字颜色、单一来源 URL/超链接和 `[1]–[6]` 的原链接目标。

## 实际验证

- 本地纠错 Golden、ZIP 等价重封装和另一套合法深蓝/深灰主题变体均为 1.0；未纠错的上游答案因 6 个无来源显示 URL 为 0。
- 原稿未修改、错图表数值、删图表、图注错配、进度比例错、漏进度条、额外 URL、错超链接目标、链接无下划线、标题/正文颜色错、丢引用、删原图、压平和答案标记均为 0；多一个调试文件为 0.96。
- fresh replay 连续 5 次均为 1.0；单题结构、输入/Golden 资产和相关单元测试均通过。
- 官方 Harbor 0.21 Oracle 容器回放 job 为 `pptarena093-oracle-climate-cascade-20260829`，trial 为 `pptarena-093-climate-multi-edit__6vSgyrP`；奖励 1.0、0 异常，容器产物 SHA256 与本地 Golden 一致。
- 7 页已渲染为接触表并逐页目视；第 4 页图表/说明、第 5 页两组图文和第 7 页 URL 无重叠裁切。`visual_qa.py` 只报告原模板的标题/正文占位框高度偏大，无溢出或重叠错误。
