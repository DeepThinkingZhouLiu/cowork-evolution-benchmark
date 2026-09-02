# 定向负例说明

`negative_cases.py` 从本题 Golden 逐份复制一个临时交付目录，再注入且只注入一个本题专属错误。覆盖八人来源字段、六条 Procurement 政策、17 条证据、224 个公式单元格、四类互斥决策、PIP 仅路由 VP、超 60 天零动作、Felix 过期、5% 费率冲突、六封邮件、六项台账更新、Slack 零披露、零执行、Word/PPT 跨文件一致性和 Office 原生结构。

负例仅写入临时目录，不修改 Golden；运行时把真实渲染替换为已验证页数，避免 1,064 个语义负例重复启动 LibreOffice。Golden、fresh replay 和正式 Harbor 回放仍执行真实 LibreOffice 渲染。
