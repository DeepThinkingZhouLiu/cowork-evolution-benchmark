# 定向负例说明

`negative_cases.py` 从本题 Golden 逐份复制一个临时交付目录，再注入且只注入一个本题专属错误。覆盖八起事故全部源字段、安全手册、10 条政策、22 条证据、538 个公式单元格、三类严重度与路由、OSHA / Legal 优先触发、精确时刻缺口、工伤职责边界、八封邮件、三份 JSON、公开 Slack 隐私、零执行、Word/PPT 跨文件一致性和 Office 原生结构。

负例仅写入临时目录，不修改 Golden；运行时把真实渲染替换为已验证页数，避免 1,495 个语义负例重复启动 LibreOffice。Golden、fresh replay 和正式 Harbor 回放仍执行真实 LibreOffice 渲染。
