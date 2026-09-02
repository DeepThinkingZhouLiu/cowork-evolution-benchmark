# AutomationBench 5130 逐题改造记录

- 上游：AutomationBench `get_task_5130()` / `hr.headcount_planning`，commit `4a8e1061254004d9dac807054eed33fad7d1ff14`，MIT；单题快照 SHA256 `a5a4232fbfa18725340a293be78bfb31740bfa04a57dd12082e31d3b3d6fc95a`。
- 已完整读取原 Query、8部门、CFO邮件、2个Slack频道、冻结时间和10条断言。`materialize_inputs.py` 只处理5130，不是通用造题器。
- 真值：Engineering 17/13/4，Sales 11/9/2；最终批准57、已填48、算术open9。Design 3/2/1但Notes含hiring freeze，因此可招聘open为0，总可招聘open为8。
- CFO邮件覆盖用户的全员分享要求。CEO仅有1封未发送草稿；leadership因政策称私有、源元数据却为公开而阻断发布；all-staff无草稿、无发布。
- `solve.py`包含5130专属推导；`office_runtime.py`只做Office渲染与序列化。Golden含12张Excel表、168个公式、3页Word、3页PPT、1封CEO草稿和3份控制JSON。
- Judge独立读取来源并核验57/48/9/8、Design freeze、八部门、CEO收件人、leadership门禁、all-staff禁止、零副作用、安全关系和12/3/3页。
- 9类负例全部0分；两次fresh replay和两次Harbor 0.21 Docker Oracle均1.0、0异常；18页目视检查通过。
- 正式难度未赋值；结构目标为L4/Expert，最终由冻结多Harness Rollout校准。
