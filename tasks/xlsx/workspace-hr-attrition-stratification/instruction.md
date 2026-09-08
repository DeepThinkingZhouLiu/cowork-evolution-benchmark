# 员工数据联表与离职比例分层复核

请从 /app/data/input_files 的8份CSV和2份字段说明，交付 /app/output/hr_review.xlsx。三份部门花名册和全员主表不是四批独立员工；请先用EmployeeNumber核对，再关联薪酬、绩效、满意度、通勤任职信息。不能按行号横向拼接、重复累计花名册，不能把MonthlyRate当MonthlyIncome，也不能把两个量表里相同数字解释成同一含义。

分析仅描述这个公开示例数据的群体差异，不预测个人、不使用性别/年龄等字段推断风险，不给出辞退名单，不把相关关系写成因果。金额沿用原表单位，不擅自声称币种。请提供可编辑公式，让管理者复核样本构成对部门离职比例的影响。

## 输出接口

第一行表头，共7个页签。

1. 联合底稿：每员工一行，按EmployeeNumber数值升序；A编号、BDepartment、CJobRole、DJobLevel、EAttrition、FMonthlyIncome、GMonthlyRate、HOverTime、IPerformanceRating、JPercentSalaryHike、KJobSatisfaction、LWorkLifeBalance、MYearsAtCompany、NDistanceFromHome、OEnvironmentSatisfaction、PRelationshipSatisfaction、QJobInvolvement、R绩效英文标签、S工作满意度英文标签。数字列为数字而非文本；标签来自两份codebook。保留源值，不造新员工。
2. 主键审计：8份CSV按文件名字典序，A文件名、B原行数、C唯一编号数、D与主表相比缺失编号数、E多出编号数、F同名可比字段冲突数。部门花名册D是主表里不在该部门花名册中的数量，不是错误；三部门并集应与主表一致。F逐单元格计数，编号本身不算；只比较双方都存在的字段。写清这一区别，不将部门内不覆盖全员误判损坏。
3. 分组统计：先Department三组，再JobRole九组，各组按字符串升序，再Attrition按No/Yes两组。A维度、B组值、C人数、D离职数、E离职比例、FMonthlyIncome均值、G工作满意度均值、H绩效均值、I低工作满意比例（K<=2）、J加班比例（H=Yes）。C:J公式；比例必须“该组离职数/该组人数”，不是全部员工分母。
4. 分层统计：部门字符串升序×JobLevel 1到5×OverTime No/Yes，30行。A部门、B职级、C加班状态、D人数、E离职数、F离职比例、G收入均值、H工作满意均值。D:H公式。若某组无员工，比例/均值留空，不填0伪装观测。展示组大小，不用小样本差异做确定性归因。
5. 职级标准化：A职级1—5、B全员该级人数、C全员该级权重、D/E/F依次为Human Resources、Research & Development、Sales部门在该职级的离职比例；B:F公式。H部门三行、I原始部门离职比例、J按共同权重标准化比例=sum(全员职级权重×本部门各职级离职比例)、K标准化减原始，I:K公式。每部门用同一组全员职级权重；这只是构成调整，不是因果估计。
6. 审查指标：A指标/B结果，B2:B11依序公式计算：唯一员工数、离职数、全员离职比例、三个部门比例的简单平均、按部门人数加权的比例、在职平均收入、离职平均收入、两者收入差（在职减离职）、PerformanceRating>=3人数、低工作满意人数。除代码本明确指出的含义，不推断薪酬单位或年度增长的因果作用。
7. 判读与边界：至少四行，A主题、B具体证据、C结论边界、D后续需核实。覆盖花名册去重、量表含义、群体构成/共同权重、非因果/小样本。不要只写“数据有偏差”；说明重复计算与错误分母/标签为什么会误导。

所有统计/标准化必须原生公式，可随底稿变化重算；不接受把答案写成常数公式。Judge会改变一人的离职状态、工资和职级，再重排底稿，测试计数、均值、分层及共同权重联动。至少一张原生图表对比部门原始/标准化比例，冻结首行、筛选，无宏和外链，不联网。无需大模型或VLM评分。
