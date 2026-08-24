---
id: Finance-018
name: Finance-018
category: Finance
grading_type: llm_judge
timeout_seconds: 2400
input_modality: document
workspace_files:
- source: input/2025往来明细-上半年.xlsx
  dest: input/2025往来明细-上半年.xlsx
- source: input/2026往来明细.xlsx
  dest: input/2026往来明细.xlsx
- source: input/2025往来明细-下半年.xlsx
  dest: input/2025往来明细-下半年.xlsx
- source: skills/financial-report-generator/SKILL.md
  dest: skills/financial-report-generator/SKILL.md
- source: skills/financial-report-generator/reference/ref-income-statement.md
  dest: skills/financial-report-generator/reference/ref-income-statement.md
- source: skills/financial-report-generator/reference/ref-balance-sheet.md
  dest: skills/financial-report-generator/reference/ref-balance-sheet.md
- source: skills/financial-report-generator/reference/ref-cash-flow-statement.md
  dest: skills/financial-report-generator/reference/ref-cash-flow-statement.md
- source: skills/budget-variance-analyzer/SKILL.md
  dest: skills/budget-variance-analyzer/SKILL.md
- source: skills/budget-variance-analyzer/assets/variance-report-template.md
  dest: skills/budget-variance-analyzer/assets/variance-report-template.md
- source: skills/budget-variance-analyzer/reference/variance-calculation-rules.md
  dest: skills/budget-variance-analyzer/reference/variance-calculation-rules.md
- source: skills/budget-variance-analyzer/reference/variance-cause-taxonomy.md
  dest: skills/budget-variance-analyzer/reference/variance-cause-taxonomy.md
- source: skills/tin-format-validator/SKILL.md
  dest: skills/tin-format-validator/SKILL.md
- source: skills/tin-format-validator/references/North-America.md
  dest: skills/tin-format-validator/references/North-America.md
- source: skills/tin-format-validator/references/Oceania.md
  dest: skills/tin-format-validator/references/Oceania.md
- source: skills/tin-format-validator/references/Asia.md
  dest: skills/tin-format-validator/references/Asia.md
- source: skills/tin-format-validator/references/Africa.md
  dest: skills/tin-format-validator/references/Africa.md
- source: skills/tin-format-validator/references/South-America.md
  dest: skills/tin-format-validator/references/South-America.md
- source: skills/tin-format-validator/references/Europe.md
  dest: skills/tin-format-validator/references/Europe.md
- source: skills/finance-data-analyzer/SKILL.md
  dest: skills/finance-data-analyzer/SKILL.md
- source: skills/finance-data-analyzer/reference/accounting-rules.md
  dest: skills/finance-data-analyzer/reference/accounting-rules.md
- source: skills/finance-data-analyzer/reference/data-cleaning-rules.md
  dest: skills/finance-data-analyzer/reference/data-cleaning-rules.md
- source: skills/invoice-contract-matcher/SKILL.md
  dest: skills/invoice-contract-matcher/SKILL.md
- source: skills/invoice-contract-matcher/reference/default-rules.md
  dest: skills/invoice-contract-matcher/reference/default-rules.md
- source: skills/iban-validator/SKILL.md
  dest: skills/iban-validator/SKILL.md
- source: skills/iban-validator/reference/iban-length-rules.md
  dest: skills/iban-validator/reference/iban-length-rules.md
- source: skills/iban-validator/reference/country-mapping-rules.md
  dest: skills/iban-validator/reference/country-mapping-rules.md
- source: skills/invoice-validator/SKILL.md
  dest: skills/invoice-validator/SKILL.md
- source: skills/invoice-validator/references/ordinary-invoice.md
  dest: skills/invoice-validator/references/ordinary-invoice.md
- source: skills/invoice-validator/references/special-invoice.md
  dest: skills/invoice-validator/references/special-invoice.md
- source: skills/data-cashflow-forecast/SKILL.md
  dest: skills/data-cashflow-forecast/SKILL.md
- source: skills/data-cashflow-forecast/reference/cashflow-exception-playbook.md
  dest: skills/data-cashflow-forecast/reference/cashflow-exception-playbook.md
- source: skills/data-cashflow-forecast/reference/cashflow-unit-normalization.md
  dest: skills/data-cashflow-forecast/reference/cashflow-unit-normalization.md
- source: skills/data-cashflow-forecast/reference/liquidity-alert-policy.md
  dest: skills/data-cashflow-forecast/reference/liquidity-alert-policy.md
- source: skills/data-cashflow-forecast/reference/cashflow-field-mapping.md
  dest: skills/data-cashflow-forecast/reference/cashflow-field-mapping.md
- source: skills/data-cashflow-forecast/reference/cashflow-forecast-rules.md
  dest: skills/data-cashflow-forecast/reference/cashflow-forecast-rules.md
solution_files:
- source: solution/rubrics.md
  dest: solution/rubrics.md
- source: solution/未收统计标准答案.csv
  dest: solution/未收统计标准答案.csv
- source: solution/未付统计标准答案.csv
  dest: solution/未付统计标准答案.csv
rubric_file_paths:
- /workspace/solution/未收统计标准答案.csv
- /workspace/input/未收统计.csv
- /workspace/solution/未付统计标准答案.csv
- /workspace/input/未付统计.csv
rubrics:
- id: '01'
  content: 输出的<file>/workspace/input/未收统计.csv</file>，与<file>/workspace/solution/未收统计标准答案.csv</file>内容一致
  weight: 1
- id: '02'
  content: 输出的<file>/workspace/input/未付统计.csv</file>，与<file>/workspace/solution/未付统计标准答案.csv</file>内容一致
  weight: 1
- id: '03'
  content: 输出的回答中，有“无僵尸账款”或其他等价表述
  weight: 1
---

## Prompt

查阅公司近两年的往来明细：
```
/workspace/input/2025往来明细-上半年.xlsx
/workspace/input/2025往来明细-下半年.xlsx
/workspace/input/2026往来明细.xlsx
```
执行以下两个任务：

1. 整理一下目前的欠款和未收情况，返回两个表格：
```
/workspace/input/未付统计.csv
/workspace/input/未收统计.csv
```
按下面的示例填写（未收统计第三列列名需调整为`未收款项`），并按客商名称升序排列，同一客商再按款项金额降序排列：

|  科目名称   | 客商名称  | 未付款项  |
|  ----  | ----  | ----  |
| 应付账款\应付资产运营业务款  | 【西瓜互娱】 | 94,339.61 |

2. 查找目前的未收款项中，同一客商同一科目在2025年上半年到2026年三个统计期间，期末余额是否存在持续未动的情况，如果连续三个统计期间期末余额无任何变化则视为“僵尸账款”，需要单独说明。此处无需生成文件，直接告诉我存在“僵尸账款”的客商名称、科目名称及金额即可。

请直接执行命令，不要问我任何问题，也不要让我做出进一步决策。如果遇到任何问题和决策点，请你自行解决。

## Grading Criteria

- [ ] [01] 输出的<file>/workspace/input/未收统计.csv</file>，与<file>/workspace/solution/未收统计标准答案.csv</file>内容一致
- [ ] [02] 输出的<file>/workspace/input/未付统计.csv</file>，与<file>/workspace/solution/未付统计标准答案.csv</file>内容一致
- [ ] [03] 输出的回答中，有“无僵尸账款”或其他等价表述
