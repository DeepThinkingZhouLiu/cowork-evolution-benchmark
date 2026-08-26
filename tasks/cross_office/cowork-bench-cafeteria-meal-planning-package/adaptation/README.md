# 逐题改造记录

本题只改造 Cowork-Bench `terminal-pw-kulinar-excel-word-gcal`，不是根据 `template_driven_package` 家族名称自编的任务。

- 上游仓库：`0717376/cowork_bench`
- 上游 commit：`d943e75bc0fc8e3b27141979300cd8cbcd1e890d`
- 上游任务路径：`tasks/finalpool/terminal-pw-kulinar-excel-word-gcal`
- 上游配方路径：`local_servers/kulinar-mcp/src/data/all_recipes.json`
- 许可：Apache-2.0

原题的工作链是：价格页 -> Kulinar 热菜配方 -> Python 成本优化 -> 三表 Excel -> Word 提案 -> 5 个 Google Calendar 备菜事件。本地仅把 HTTP、MCP 和日历副作用静态化为 HTML、完整 JSON 和 ICS，保留上游的时间、文件名、表名和关键业务约束。

逐文件阅读后发现，上游 `groundtruth_workspace` 仍是旧版英文/美元中餐数据，与当前俄文 Query、卢布价格页和俄罗斯配方库冲突，因此不能作为本题 Golden。本地 Judge 直接从 22 行价格和 50 条配方重算。

上游还存在两个必须显式处理的歧义：

- 报价单位与配方用量无完整换算关系；本题因此评测“报价篮子成本”，不允许虚构重量、包装量或未报价单价。
- PDF 要求每份至少 20g 蛋白质，但任何输入都没有营养成分；本题必须披露“无法验证”，不得伪造合规结论。

Golden 选中的 5 道菜不是手工挑选，而是按指令中可回放的“至少 3 个匹配原料 -> 报价篮子成本升序 -> Recipe_ID 升序”规则从上游输入计算得到。
