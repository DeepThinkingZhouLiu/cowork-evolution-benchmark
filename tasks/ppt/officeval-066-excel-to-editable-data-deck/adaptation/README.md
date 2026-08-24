# 改造说明

本题逐题改造自 OmegaUse-OfficeVal 的 `officeval_066`，精确来源版本为 `ffbeecb8752447c8e40b594a0eeb1db7236ecb36`。

- `upstream-task.json`：上游真实 Query 与两个输入文件地址，原样保留。
- `upstream-rubric.json`：上游逐项评分要求，原样保留。
- `../tests/upstream_verifier.py`：上游逐题 Verifier，原样保留并由本题 Judge 直接调用。
- `LICENSE-OmegaUse-OfficeVal`：上游 Apache-2.0 许可证。

本 Benchmark 没有自编 Query、输入数据或 Golden 事实。改造只增加固定输出文件名、离线执行边界、前 4 页保留性检查、可编辑性检查、实际渲染检查和 Harbor 所需元数据。参考解法逐格读取真实 Excel，并在真实 PPT 中追加两页；它不是批量制题脚本。
