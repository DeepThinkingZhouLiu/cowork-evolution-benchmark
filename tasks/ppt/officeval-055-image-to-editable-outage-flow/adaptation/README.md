# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_055`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 和 PNG 地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `upstream-verifier.py`：上游逐题 Verifier，原样保留；`../tests/` 中亦保存了可离线执行的副本。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query 或输入数据。真实输入 `website_outage_editable_flow.png` 来自上游公开数据，SHA256 为 `80bca602124e0cd54203937d2d46817062e7c61849cc05c911e64345be1c6add`。本地只去掉了上游文件名在扩展名前的一个空格，文件内容未变。

Golden 是针对这一道真实题的逐对象改造结果。我先人工查看原图，核对五个阶段的编号、图标、标题、正文、箭头、情绪折线和底部主题，再按上游 Query 把第二阶段改为 `check`，并替换 `finish` 正文。交付物中没有嵌入原图；每个文字框、形状、图标部件、箭头和折线都可单独编辑。

`solve.py` 只复现本题的人工改造结果：它先校验这一个真实输入的 SHA256，然后生成本题固定的可编辑交付物，不是通用制题器，也不批量 mock 其他题。

上游 Verifier 有两个已确认的实现问题。第一，它在计数四条主 action 箭头时没有排除左上角图例中的黑色箭头，因此会得到五条并误判；第二，它仅用 OOXML 连接线的正数 `ext.cy` 判断上下方向，却忽略了真正表示向上折线的 `flipV`。原始 Verifier 保留不改；本地 Judge 排除图例区域，并按 `flipV` 判断视觉方向。

参考交付物已通过 LibreOffice 实际渲染、人工视觉检查、Golden 正例、错误文本、缺失阶段、缺失箭头、缺失 emotion 折线、空白 PPTX 和整页截图等负例验收。
