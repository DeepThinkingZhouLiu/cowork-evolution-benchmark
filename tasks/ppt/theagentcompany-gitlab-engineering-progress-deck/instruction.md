请基于下列冻结输入，为 CTO Sarah Johnson 制作一份 GitLab 工程进展演示文稿：

- `/app/input_files/Original_Request.md`
- `/app/input_files/GitLab_Repository_Snapshot.json`

仅交付 `/app/output/Engineering_Progress.pptx`。演示文稿必须满足：

- 恰好 15 页，16:9；第 1 页标题必须是 `Overview of Our Engineering Progress`。
- 封面准确总结冻结快照：14 个仓库、10 个有活动记录的仓库、4,990 条 Issue、794 条 Merge Request。
- 第 2–15 页一仓库一页，严格按冻结快照的 `order` 顺序排列；每页明确展示仓库名、描述、Issue 总数和 Merge Request 总数。
- 计数口径是所有状态：Issue 包含 open + closed，Merge Request 包含 open + merged/closed。这些数字表示活动量，不能宣称为当前 backlog 或工程质量。
- 冻结快照不变，所以所有数值必须精确，不采用上游在线 evaluator 的 `+/-5` 容差。
- `api-server`、`copilot-arena-server` 和 `raft` 的上游描述为空，页面必须显示 `No description set in frozen GitLab snapshot`，不得自行补写业务描述。
- 每页保留来源，演示文稿中所有文本、表格、图形和数据图形必须原生可编辑，不得将页面压成整页图片。
- 成品要能在 LibreOffice Impress 中打开和渲染，不得包含宏，不得出现重叠、裁切或超出页面。
