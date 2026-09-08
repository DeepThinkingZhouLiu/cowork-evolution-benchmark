# 升级七道 Excel 满分题

- 用途：从 Workspace-Bench 逐题选择有完整来源的任务与输入，根本性重做七道简单满分题；不以修改评分标准代替任务升级。
- 分支：`lz/data-xlsx-hardening-7`
- 基线：`main@3c11d86335159d2890c84cd3135bdab35511b044`
- 状态：eval（七题构造/作者侧验收完成，待真实Agent重测和提交发布）
- 预期产出：七个独立 Harbor 题包、逐题上下文阅读记录、专属 Golden/Rubric/Judge、正负例与 Oracle 回放证据。
- 验证基线：2026-09-08，任务结构与资产校验单测 8/8 通过。
- 当前产出：七题替换完成，240题资产通过，7/7官方Harbor Oracle通过；独立正负例/重复阅卷通过。详见 docs/xlsx-hardening-7-progress.md。
- 下一步：七题已提交，v0.4已生成并校验；执行Claude Code/Sonnet 5七题重测，合并53题历史证据后报告四档试分档。尚未合入main。
