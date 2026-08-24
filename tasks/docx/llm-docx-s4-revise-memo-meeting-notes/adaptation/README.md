这道题逐题改造自 `LLM-DOCX-Editing` 的 S4，不是本项目编写的业务故事。

- 原始输入来自上游 `packages/agent/test/fixtures/legal-memo.docx`。
- 原始用户要求来自上游 `packages/agent/test/review-test.ts`，在本题 `instruction.md` 中逐字保留。
- 上游把这道题定义为把 2026 年 3 月 12 日客户会议记录合并进既有法律备忘录；这与本项目 `merge_compare` 家族的“把另一版本的信息有选择地合并到主文档，并保护未修改内容”一致。
- 来源仓库把它作为公开能力测试编写，因此 provenance 明确标记为合成样题；它不是自然产生的客户工单，也没有被包装成真实客户资料。
- Golden 直接来自上游 `docs/showcase/results.json` 中该题 `passed=true` 的公开编辑结果。解码后 SHA256 为 `6f623d6702e418f837f1628ccf553fd1895f4672236cda53da174edd8356e6ff`，本项目没有另编一份标准答案。
- 逐段比对确认：上游结果修改 7 段、未新增或删除段落；摘要、时间、预算、Option B、比较分析和建议均已更新，背景事实及 Option A 章节保留。
- 同仓库的 `docs/claude-code-outputs/s4-revise-memo-with-meeting-notes.docx` 把多个目标段落清空，实际不满足任务要求，因此没有用作 Golden。
- 输入与 Golden 均由 LibreOffice 实际渲染为两页 Letter 文档并逐页查看，没有空页、截断或截图替代正文。
- 本项目增加的 Judge 不要求答案逐字复制 Golden，而是分别核对新决策的一致传播、旧决策清除、关键事实保留、章节结构、可编辑性和真实渲染。

正式发布前仍需让至少一个非参考 Agent Harness 独立作答，以估计题目难度与区分度。
