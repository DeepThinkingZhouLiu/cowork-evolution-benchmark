# 定向负例说明

`negative_cases.py` 从本题 Golden 逐份复制一个临时交付目录，再注入且只注入一个本题专属错误。覆盖八岗来源字段、历史发布邮件、七条任务规则、20 条证据、422 个公式单元格、五类互斥决策、90 天严格边界、冻结/暂停/已招满零动作、stale VP 批准前不重发、rejected-pool 候选人零虚构与零联系、六封邮件、三份 JSON、Slack 零发布、零执行、Word/PPT 跨文件一致性和 Office 原生结构。

负例仅写入临时目录，不修改 Golden；运行时把真实渲染替换为已验证页数，避免 1,187 个语义负例重复启动 LibreOffice。Golden、fresh replay 和正式 Harbor 回放仍执行真实 LibreOffice 渲染。
