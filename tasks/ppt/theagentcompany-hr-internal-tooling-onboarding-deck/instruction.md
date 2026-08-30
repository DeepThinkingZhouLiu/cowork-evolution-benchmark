请基于以下冻结输入，制作 The Agent Company 新员工内部工具入职培训包：

- `/app/input_files/TAC_overview.md`
- `/app/input_files/Personnell_File.odt`
- `/app/input_files/Internal_Tooling_Source_Snapshot.json`
- `/app/input_files/Task_Delivery_Context.json`

主交付物必须是 7 页、16:9 的 `Internal_Tooling_Training.pptx`，页面顺序和标题为：

- `Internal Tooling Training`：面向新员工的业务风格封面。
- `TAC Overview`：准确总结 mission、vision 和 6 类主要产品/服务。
- `RocketChat`：包含 URL，完整介绍冻结快照中的 15 个频道和用途，并明确 `#general` 与 `#kudos` 是 default channels。
- `Gitlab`：包含 URL，分别说明 `code collaboration`、`project management` 和 `issue tracking` 的使用方式。
- `Gitlab Repos`：列出并简介快照中全部 14 个仓库；描述不得超出 `evidence` 边界。
- `Owncloud`：包含 URL，介绍冻结镜像 `/Documents` 下实际 9 个顶层目录。必须暴露上游 evaluator 遗漏 `Marketing`、`Q1 Planning and Allocation` 且将 `Human Resources Team` 误写为 `Human Resource Team` 的语义差异，并明确以冻结镜像为准。
- `Plane`：包含 URL，介绍 `projects -> modules -> issues -> sprint cycles -> assignees -> analytics` 工作流，以及 `Total Open Tasks`、`Pending Issues`、`Most Issues Created`、`Most Issues Closed` 4 个分析指标。

原任务需要上传 OwnCloud 并给 HR 经理发 RocketChat。本 Harbor 适配为无网络、可重放设计，不得伪造已发送成功；请另外交付 `Delivery_Record.json`，准确记录：

- 原始 OwnCloud 路径 `/Documents/Human Resources Team/Internal_Tooling_Training.pptx`。
- 收件人 `Chen Xinyi`，通道 `RocketChat direct message`。
- 目的 `review and inclusion in the onboarding materials`。
- 离线共享引用 `owncloud:/Documents/Human Resources Team/Internal_Tooling_Training.pptx`。
- `network_side_effect` 必须是 `not_executed_offline`。

演示中的所有文本和图形必须原生可编辑，无重叠、无裁切、无整页压平，并在每页保留来源。

请仅在 `/app/output` 中写入：

- `/app/output/Internal_Tooling_Training.pptx`
- `/app/output/Delivery_Record.json`
