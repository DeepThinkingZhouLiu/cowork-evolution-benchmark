# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_041`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 和五个输入文件地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `../tests/upstream_verifier.py`：上游逐题 Verifier，原样保留。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query、图片、工作簿或源演示文稿。五个输入文件均从上游 `officeval_041` 下载，文件名、字节数和 SHA256 记录在 `metadata/input-manifest.json`。源演示文稿共 23 页，原始逐页备注完整保留；其中第 6 页是目录，原题要求把它移到第 2 页。

Golden 是对这一道真实题的逐页参考交付，不是批量 mock：先人工阅读 23 页源稿、全部备注、三张现场图片和 Excel 的 `Sheet1!A1:D6`，再针对这一题编写专用 `solve.py`。它只完成原备注所要求的目录重排、字体和字号、三图插入、Excel 表格替换、视频自动播放、四页图片动画、文本扩充、颜色和版式调整，不服务其他题。

上游 Verifier 的“第19页Excel表格具体数据”存在一处与上游真实输入冲突的错误。真实工作簿的 `Sheet1!D2` 是“约12,800件/小时”，`Sheet1!D4` 是“模块抽拉式维护，关键部位快速接入”；上游 Verifier 却错误地假定方案 C 的这两个单元格为空，并禁止出现额外“件/小时”值。Golden 忠实保留真实工作簿的全部 24 个单元格，因此上游 Verifier 为 65/66。本地 Judge 继续复用其余上游规则，并将这一项修正为“与真实工作簿 A1:D6 逐格一致”。

参考交付物已经由 LibreOffice 实际渲染为 23 页，并人工查看全部页面。重点修改页 P2、P4、P9、P11、P13、P14、P15、P16、P17、P18、P19、P21、P22、P23 均检查了文字可读性、边界、遮挡和表格完整性。
