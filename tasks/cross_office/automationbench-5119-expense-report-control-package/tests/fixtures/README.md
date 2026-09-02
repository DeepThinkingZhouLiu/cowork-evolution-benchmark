# 定向负例说明

`negative_cases.py` 从本题 Golden 逐份复制一个临时交付目录，再注入且只注入一个本题专属错误。覆盖输入事实、政策全文、19 条证据、22 列公式、八人逐行决策、三个动作队列、Elena 静默保持、七封邮件、七项台账更新、Slack/执行零副作用、Word/PPT 跨文件一致性和 Office 原生结构。

负例仅写入临时目录，不修改 Golden；运行时把真实渲染替换为已验证页数，避免数百个语义负例重复启动 LibreOffice。Golden、fresh replay 和正式 Harbor 回放仍执行真实 LibreOffice 渲染。
