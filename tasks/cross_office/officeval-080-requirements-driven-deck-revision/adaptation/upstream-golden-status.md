核验结论：没有发现可公开下载的 `officeval_080` 上游任务级 Golden。

- 核验版本：Hugging Face commit `cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5`。
- `task_files/officeval_080/` 仅包含 `公开课课件.pptx` 与 `要求.docx`。
- 数据卡汇总表声称整个数据集有 115 个 Golden artifacts，但同一固定版本展示的公开目录结构只列出输入、任务、rubric 和数据表，没有 Golden 目录或 `officeval_080` Golden 路径。
- GitHub commit `ffbeecb8752447c8e40b594a0eeb1db7236ecb36` 提供本题 task、rubric 和 verifier，也没有 task-level Golden 或 preprocess。

所以本题 `solution/` 中的 PPTX、响应表和解析计划是对两份真实上游输入逐条人工完成、再由 Judge 复核的本地参考解。它们不是“复制来的上游 Golden”，也不把本地判断冒充上游发布内容。
