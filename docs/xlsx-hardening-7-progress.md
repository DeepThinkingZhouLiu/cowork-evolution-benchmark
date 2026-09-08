# 七道 Excel 满分题逐题升级记录

更新时间：2026-09-09。分支 lz/data-xlsx-hardening-7，基于 main@3c11d86335159d2890c84cd3135bdab35511b044；本轮没有commit、push或合并主线。

## 已完成的替换

| 旧题                                      | 新题                                   | 内容升级                                      |
| ----------------------------------------- | -------------------------------------- | --------------------------------------------- |
| officebench-1-9-total-salary                | workspace-payroll-reconciliation        | 四份工资/报告材料的版本、人数、口径及情景核对  |
| officebench-1-10-lowest-budget-add-highest  | workspace-admin-budget-period-audit     | 跨期间预算、万元换算、共享预算、现金窗口审计    |
| officebench-1-9-total-shopping-weight       | workspace-canteen-stock-procurement     | 库存勾稽、按品取整采购、调价情景和合规证据边界 |
| officebench-1-14-score-column-swap          | workspace-trade-series-audit            | 五文件十年贸易、年度/累计差额、同比与加权份额  |
| officebench-1-10-highest-salary-duplicate   | workspace-annual-payroll-ledger-audit    | 96条月薪、年度勾稽、重复扣款情景和波动分析      |
| officebench-1-6-attendance                 | workspace-hr-attrition-stratification   | 1470员工多源关联、30层群体统计和共同职级标准化  |
| officebench-1-9-score-average              | workspace-commerce-metric-reconciliation | 天猫周报、单品净化、比例分母与历史证据复核     |

每题由主代理读实际材料后分别设计 instruction、Golden、solver、Rubric 和 Judge，没有通用题目生成器。下载、复制、哈希、测试启动、元数据一致性检查属于机械操作。输入均来自 Workspace-Bench 固定提交的公开上游模拟办公素材，共32个原始文件，字节不改；不是代理自造输入，不宣称真实企业私有数据。具体来源、阅读范围和设计在各题 adaptation/README.md。

全部题使用v0.2既有家族/能力标签，正式难度保持unassigned。结构复杂度预估不等于实测Hard/Expert；旧七题满分记录保留，不能用于新题。

## 验证结果

| 新题简称 | 输入文件 | Fresh/Golden | 重复评分          | 最终官方Harbor Oracle  | 常数伪公式 |
| -------- | -------: | ------------ | ----------------- | ---------------------- | ---------: |
| 薪酬核对 |        4 | 1.0 / 1.0    | 5次均1.0          | 1/1，1.0，0异常        |       0.79 |
| 预算审计 |        4 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.78 |
| 库存采购 |        4 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.77 |
| 贸易审计 |        5 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.75 |
| 年度工资 |        1 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.76 |
| HR分层   |       10 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.75 |
| 网店复核 |        4 | 1.0 / 1.0    | 3次均1.0          | 1/1，1.0，0异常        |       0.75 |

七题均覆盖空输出、损坏文件、空白工作簿、未完成业务的原样/部分答案，以及各题专属口径错误、错误输入和公式变更。空/损坏/空白均0；HR只交正确关联底稿仍有0.197143部分分，不是假装没有完成任何内容。少量局部错误只扣相关项，不为压低分数任意全局归零。

- 各题 adaptation/validation-summary.json 保留细分分数；validate_reference.py 可重跑专属正负例。
- 公式测试会重算并修改真实业务底稿，检查改变后的输出；固定答案伪装成公式不能通过。
- Judge独立读取输入而非Golden，不调用solver。文字说明使用有限关键词/结构检查，不能宣称完整语义评分；不新增VLM。
- Oracle只验证参考解及隔离运行链，不是Claude/OpenCode/OpenHands性能。
- Golden均完成PDF导出和指定关键页抽查，阅读页码在各题记录中；没有宣称全部页面VLM验收。
- 第七题首次构建在Debian下载中断，停止后复用已验证基础层，最终回放正常。临时清除代理并用host构建网络的诊断有下载进展，但在导出镜像前结束；不能由此断言代理根因，更不归因成API Key失效，不把中断记为任务0分。
- 全仓四个正式领域各60题，资产校验240/240；七题Rubric Schema、metadata/task.toml一致性、输入哈希通过；任务相关单测8/8。
- 扩大Schema抽查时发现未改动的 workspace-bank-ledger 仍使用旧维度 content_fidelity，不在本轮七题范围；未宣称全仓Rubric均通过最新Schema。

## 分区与发布边界

旧 -> 新 ID完整映射见 xlsx-hardening-7-replacements.json。新题逐一继承被替换题的分区，经当前内容审计：
- PPT、DOCX、XLSX仍各30/10/20，Cross-Office仍0/0/60；总90/30/120。
- 240个任务、182个来源/共享文件连通组件，跨分区组件0；具体见 xlsx-hardening-7-integrity.json。
- 现有检测覆盖来源ID、泄漏组和相同文件SHA256，不声称能自动排除所有语义近重复。
- 本轮未修改历史v0.3，也未伪造v0.4：冻结器绑定HEAD:tasks，而任务尚未commit。当前是已验证工作树，不是已发布Release；不能直接拿旧v0.3锁加载新任务路径。
- 后续得到提交授权后，先提交七题，再按映射生成v0.4、复核内容锁/分区/泄漏并提交Release；随后才能发布或合并。其余233题、旧结果不应被覆盖。
- 七道新题需重新跑真实Agent，不能只重评旧产物。尚未启动模型评测或调用Provider。

本地完整日志：.coworkevo/xlsx-hardening-7/。旧题从当前升级工作树移除，原输入/答案可从main与Git历史找回。主线未被本轮改动。
