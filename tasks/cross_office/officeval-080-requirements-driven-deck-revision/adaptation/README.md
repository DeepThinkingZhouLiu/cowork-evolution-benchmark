本目录保存 OmegaUse-OfficeVal `officeval_080` 的完整、可核对上游材料。本题不是凭空编题，也没有批量生成输入。

来源固定如下：

- GitHub 仓库：`https://github.com/baidu-frontier-research/OmegaUse-OfficeVal.git`
- GitHub commit：`ffbeecb8752447c8e40b594a0eeb1db7236ecb36`
- Hugging Face 数据 commit：`cd6ba6d8fb83b3fb551e24eebc20e1fb0bd154a5`
- 上游任务：`officeval_080`
- 原 Query：`这份PPT是需要修改的，文档里有修改要求，按照修改要求修改一下PPT`
- 许可：Apache-2.0；本目录保留上游 `LICENSE`、`NOTICE` 和第三方许可说明。

上游材料与本地文件对应关系：

- `upstream-task-zh.json` / `upstream-task-en.json`：原始中英文任务定义。
- `upstream-rubric-zh.json` / `upstream-rubric-en.json`：原始中英文评分标准。
- `upstream-verifier.py`：原始逐题 evaluator，未做修改。
- `HF-DATASET-README.md`：固定到上述 Hugging Face commit 的数据卡，用于核对许可、构造流程与隐私声明。
- `LICENSE-OmegaUse-OfficeVal`、`NOTICE-OmegaUse-OfficeVal`、`THIRD_PARTY_LICENSES-OmegaUse-OfficeVal.md`：再发布许可材料。

两份业务输入 `公开课课件.pptx` 和 `要求.docx` 保存在 `data/input_files/`，与上游下载文件逐字节一致。Word 中四张图片也直接来自该上游文件；本次适配没有联网补图，没有编写新的课程事实或业务记录。

数据卡明确说明：上游先对真实工作需求做隐私保护改写，再用 LLM 辅助重建输入并由人工修订。因此本题元数据必须诚实标记 `human_rewritten: true`、`synthetic: true`。这里的 `synthetic` 指上游重建输入，不代表本次 Harbor 适配生成了虚构输入。

本地适配只增加三项技术性输出：逐条修订响应表、可独立运行的要求解析器和解析计划。它们让 Judge 能从输入动态重算真值，并验证内容变异后处理器是否随输入变化；没有改变原 Query 的业务意图。

上游没有为单题发布 preprocess 文件。上游 GitHub 的任务级材料只有 task、rubric 和 verifier；Hugging Face 固定 commit 的 `task_files/officeval_080/` 只有两份输入。数据卡虽然汇总写有 115 个 Golden artifacts，但该固定版本的公开目录结构未包含 `officeval_080` 的任务级 Golden，详见 `upstream-golden-status.md`。因此本地 `solution/` 是基于真实输入逐条人工完成的可复核参考解，不冒充上游 Golden。

上游 verifier 有两个明确覆盖缺口：其 rubric 包含“第16页右下角图片删除”和“三处以上文字重叠或越界扣分”，但主流程的 `plus_rules` 没有前者，也没有任何扣分规则。本地 Judge 补齐这两项，并额外检查输入溯源、响应表一致性、动态处理器和变异回放。
