# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_076`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 与 PPTX 地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `../tests/upstream_verifier.py`：上游逐题 Verifier，原样保留。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query 或输入数据。输入 `模版.pptx` 的 SHA256 为 `fae5afedb2b42e6b0c4d729f0e99682fbc041b95cd0dfbcb04b3c6a3c0cd0a8b`，与上游数据文件一致。

Golden 为这一道真实题逐页制作的参考交付物：先人工检查原始 Query 中九个板块、十八个二级标题和正文，再检查真实模板的两页结构、整页背景图、文本对象和主题风格，最后制作 21 页封面、目录、九组章节页/正文页和模板尾页。每张正文页的标题与正文都保持可编辑，并有一次单击后自动连续播放的动画主序列。

`solve.py` 只服务这一道题，它校验真实模板哈希，按本题的固定大纲制作交付物，不是通用或批量制题脚本。

上游 Verifier 有一处与原始 Query 冲突的口径：它要求 `3.1 竞争格局变化` 的正文必须包含“竞争”、“格局”、“变化”、“对手”或“份额”之一，但上游 Query 给出的真实正文是以“艺术漆赛道已进入品牌集中发力阶段”开头，完全不含这些关键词。为了不篡改真实原文，Golden 保留 Query 逐字内容；本地 Judge 对该段改为“精确原文 + 位置 + 字号”检查，其余上游规则继续复用。

参考交付物已经 LibreOffice 实际渲染为 21 页 PDF，并完成全部 21 页的人工视觉检查。
