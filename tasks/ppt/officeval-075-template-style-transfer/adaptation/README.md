# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_075`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 与两个 PPTX 地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `../tests/upstream_verifier.py`：上游逐题 Verifier，原样保留。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query 或输入数据。Golden 是为这两个真实输入逐页制作的参考交付物：先人工检查 5 页参考 Deck 的页面角色、背景和主配色，再检查 39 页待改课件的页序、文字、图片和对象命名，最后逐类处理封面、目录、章节、政策、项目明细、注意事项和结束页。`solve.py` 只服务这一道题，不是通用或批量制题脚本。

上游 Verifier 有三处与真实源文件或原始 Query 冲突的口径：

- H 把源文件目录副标题的中文双引号 `“”` 误写成 `‘’`。
- J 假定第 4–7 页三个小标题必须在同一横行，但第 5–7 页真实源版式是竖向三卡。
- L 把源文件的“不得同时收费”误写成“不同时收费”。

本地 Judge 继续直接复用上游其余规则，只根据真实源文件修正这三处，并新增业务文字逐页保留、图片 SHA256 保留、新增截图检测和 LibreOffice 实际渲染检查。人工渲染抽查还发现，源第 4 页和第 9–37 页底部说明框高度过小，文字在 LibreOffice 中几乎不可读；Golden 在不改原文的前提下修正了这些真实页面，Judge 也逐页检查它们。原始 Verifier 依然留在任务中，便于审计者复核差异。
