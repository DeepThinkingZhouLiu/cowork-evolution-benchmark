# 改造说明

本题逐题改造自 OmegaUse-OfficeVal officeval_003，精确代码版本为 ffbeecb8752447c8e40b594a0eeb1db7236ecb36，数据版本为 cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5。

- 原始 Query、Rubric 和 Verifier 按中英文版本原样保留。
- 两个真实输入文件逐字节保留，并同时进入 Harbor 环境构建上下文。
- 我逐行读取了来源 DOCX：原表格共有 108 行，其中 95 行有文字、13 行为空；Golden 删除空行，保留 95 条原始记录及顶部全部字段。
- solve.py 只复现本题的人工改造结果，会先校验真实 DOCX 的 SHA256，不是通用制题或批量 mock 脚本。
- Judge 不以 Golden 哈希作为答案，而是独立检查非空行数量、横线、等距行高、顶部字段、关键内容、可编辑性和 LibreOffice 渲染结果。

上游数据卡将此数据标注为经隐私改写、模型辅助重建并经人工修订，因此 metadata 中诚实记录 human_rewritten=true、synthetic=true。
