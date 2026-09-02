# 定向负例说明

`negative_cases.py` 从本题 Golden 逐份复制一个临时交付目录，再注入且只注入一个本题专属错误。覆盖八人来源字段、两封邮件、16 条证据、192 个公式单元格、四类分流、Naomi 供应商覆盖、Failed 候选人禁发、Pending 零动作、六封邮件、六项台账更新、Slack 最小披露、零执行、Word/PPT 跨文件一致性和 Office 原生结构。

负例仅写入临时目录，不修改 Golden；运行时把真实渲染替换为已验证页数，避免 931 个语义负例重复启动 LibreOffice。Golden、fresh replay 和正式 Harbor 回放仍执行真实 LibreOffice 渲染。
