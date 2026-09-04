# 补全工资预测公式

请在 `salary_model_input.xlsx` 的 `P10 SFT` 工作表中补全就业记录行的工资预测公式。对有级别和步级数据的记录（第 8–15 行），用 `AnnIncrement` 与 `CBAIncrement` 表、`Classification`、`Step`、`Step Anniversary Date` 和 CBA `Effective Date` 计算：

- M 列：周年日所在月份的按日比例月薪；
- O 列：CBA 生效后的年度工资；若周年日在 CBA 生效日前，先前进一个步级；
- Q 列：六月过渡月按 21/31 与 10/31 加权的月薪。

请只使用可重算公式，不要硬编码结果；保留其他工作表和原有内容。将文件保存为 `/app/output/salary_model_completed.xlsx`。
