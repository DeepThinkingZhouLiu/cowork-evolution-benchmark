这道题逐题改造自 `LLM-DOCX-Editing` 的 F6，不是本项目编写的业务故事。

- 输入 `contract-with-redlines.docx`、Query 和测试期望均来自上游精确 commit；来源是公开合成能力样题，不是自然产生的客户合同。
- 输入 OOXML 中实际存在 5 个插入节点和 4 个删除节点，不是普通文本模拟。其中 4 组是行内替换，另有一整段价格调整条款包在段落级插入节点中。
- 接受视图应保留 `Global Manufacturing Corp`、45 天交付、24 个月质保、每年一次价格调整、至少 90 天书面通知、3% 或 CPI 上限以及 90 天终止通知；旧供应商和旧期限必须删除。
- Golden 直接来自上游 `docs/showcase/results.json` 中 `passed=true` 的结果，SHA256 为 `17343651fe2dc4837e4ff8c8af7132a7ec8ffe8234b2a0a0f8e11e47a063e8a2`。另一份公开 Claude 输出的字节哈希不同，但 10 个接受后段落与结构均通过同一 Judge。
- 输入与 Golden 均实际渲染为 1 页 Letter。输入显示红线旧值和新值；Golden 只显示接受后的清洁文本，未见裁切。
- 本项目 Judge 除上游关键词外，还检查完整价格调整条款、10 段接受视图逐段一致、全部旧值清除、OOXML 修订节点为零以及真实渲染。
- Golden、参考解法 replay 和独立 Claude 输出均为 `1.0`；原始红线文件、遗漏终止期替换、可见文本正确但底层残留 `w:ins` 节点和空文件均被拒绝。

正式发布前仍需由至少一个非参考 Agent Harness 独立作答，以估计题目难度与区分度。
