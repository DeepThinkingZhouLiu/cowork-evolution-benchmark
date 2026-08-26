本题一对一改造自 OmegaUse-OfficeVal `officeval_059`。

- 原始 Query、中英文 Rubric、上游 Verifier、4 页 Word 申报书和 18 页 PPT 已原样保留；输入 SHA256 记录在 `metadata/input-manifest.json`。
- 上游没有公开单题 Golden。本题实际阅读 Word 全部 4 页、定位最后一行右侧单元格，并检查 PPT 首页及 18 页结构后，加入了真实 OOXML 外部超链接。
- 链接目标严格为同目录相对文件名 `青梧谷灵MR共创计划.pptx`；随附 PPT 与输入逐字节一致。
- Word 只有 `word/document.xml` 和 `word/_rels/document.xml.rels` 两个包成员发生变化，原正文、表格、页眉、页码和边框均保留。
- 上游 Rubric/Verifier 错把真实项目“青梧谷灵”写成“海潮艺境”，并要求不存在的 `海潮艺境MR共创计划.pptx`。本地 Judge 以真实输入文件、原 Query 和可点击相对链接为真值，因此没有迎合这处错误。

Golden 已经 LibreOffice 渲染为 4 页 PDF 并逐页目视检查；本地独立 Judge 得分为 1.0。上游 Verifier 因上述名称冲突只能得到 4/19，此差异已如实记录。
