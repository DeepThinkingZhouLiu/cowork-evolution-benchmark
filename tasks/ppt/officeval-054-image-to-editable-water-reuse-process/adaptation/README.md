本题一对一改造自 OmegaUse-OfficeVal `officeval_054`。

- 原始 Query、中英文 Rubric、上游 Verifier 和 `路径图.png` 已原样保留，输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题在实际阅读参考图后，逐对象重建标题、10 个处理步骤、4 个辅助说明、异常回流、绿化与冲洗回用支路和底部四个保障模块。
- 上游中文 Rubric 把第 7 和第 8 步都写成“超滤膜组件”，但原始参考图的第 8 步清楚为“消毒接触池”。本题以原图为事实真值，并在 instruction 和 Judge 中显式固定该纠错。
- 流程图用显式底板固定为 12.44×7.00 cm，左上角为 (10.70, 6.00) cm，在 16:9 页面上居中。
- 交付页面共含 114 个原生可编辑对象，不含任何位图。

Golden 已经 LibreOffice 渲染为 1 页 PDF 并目视检查；本地 Judge 得分为 1.0。
