这道题逐题改造自 `LLM-DOCX-Editing` 的 D3，不是本项目编写的业务故事。

- 输入 `cap-conditional-edit.docx`、原始 Query、能力标签和测试期望均来自上游精确 commit。
- 来源仓库把它作为公开能力测试编写，因此 provenance 明确标记为合成公开样题；它不是自然产生的并购文件。
- 上游 Query 只允许改写 `Closing Conditions`，并要求五项成交条件全部保留；本项目没有扩大任务范围。
- Golden 直接来自上游 `docs/showcase/results.json` 中 `passed=true` 的公开编辑结果，SHA256 为 `041747580bbe6a45d3f632bc2fc972c70589fc5cee3a8d0a5035fd1e6fd0c232`。它只修改第 17 段，另外 21 段原样保持。
- 逐项核对确认五项条件分别是：交割时陈述仍准确、取得政府反垄断批准、业务没有重大负面变化、取得 Schedule 3 中第三方批准、买方落实融资。
- 同仓库一个 Claude 短输出删除了金额、买方陈述、排他条款等非目标内容，还保留了修订痕迹，因此没有用作 Golden；另一个较长输出虽语义可用，也不是本题 canonical passed 结果。
- 输入和 Golden 均实际渲染为 1 页 Letter，已逐页查看，没有截断或截图正文。
- 本项目 Judge 比上游最小关键词检查更严格：分别核对五项条件、两处术语解释、目标节外零改动、章节顺序和实际渲染，但不要求逐字复制 Golden。
- `solve.sh` 从原始 fixture 重新生成的 replay 与 Golden 的 22 个段落逐字一致，二者 Judge 均为 `1.0`；另一份措辞不同但语义完整的上游长输出也得到 `1.0`，证明 Judge 不要求答案照抄 Golden。
- 原始未作答文件、缺少一项成交条件、越界修改购买价、上游有缺损的 Claude 短输出、空白 DOCX 五类负例均被拒绝；所有 Judge 结果通过统一结果 Schema。

正式发布前仍需由至少一个非参考 Agent Harness 独立作答，以估计题目难度与区分度。
