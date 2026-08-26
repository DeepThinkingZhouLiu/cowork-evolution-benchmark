---
id: Legal-020
name: Legal-020
category: Legal
grading_type: llm_judge
timeout_seconds: 2400
input_modality: document
workspace_files:
- source: input/授权委托书风控校验规则.md
  dest: input/授权委托书风控校验规则.md
- source: input/委托书/委托书3.docx
  dest: input/委托书/委托书3.docx
- source: input/委托书/委托书2.docx
  dest: input/委托书/委托书2.docx
- source: input/委托书/委托书5.docx
  dest: input/委托书/委托书5.docx
- source: input/委托书/委托书4.docx
  dest: input/委托书/委托书4.docx
- source: input/委托书/委托书6.docx
  dest: input/委托书/委托书6.docx
- source: input/委托书/委托书1.docx
  dest: input/委托书/委托书1.docx
- source: skills/ad-compliance-checker/SKILL.md
  dest: skills/ad-compliance-checker/SKILL.md
- source: skills/ad-compliance-checker/reference/ad-violation-keywords.md
  dest: skills/ad-compliance-checker/reference/ad-violation-keywords.md
- source: skills/contract-template-filler/SKILL.md
  dest: skills/contract-template-filler/SKILL.md
- source: skills/legal-risk-graph-builder/SKILL.md
  dest: skills/legal-risk-graph-builder/SKILL.md
- source: skills/legal-contract-processor/SKILL.md
  dest: skills/legal-contract-processor/SKILL.md
- source: skills/legal-contract-processor/references/capability-2-comparison.md
  dest: skills/legal-contract-processor/references/capability-2-comparison.md
- source: skills/legal-contract-processor/references/capability-1-field-extraction.md
  dest: skills/legal-contract-processor/references/capability-1-field-extraction.md
- source: skills/legal-contract-processor/references/capability-4-structured-export.md
  dest: skills/legal-contract-processor/references/capability-4-structured-export.md
- source: skills/legal-contract-processor/references/capability-3-summary.md
  dest: skills/legal-contract-processor/references/capability-3-summary.md
- source: skills/contract-risk-review/SKILL.md
  dest: skills/contract-risk-review/SKILL.md
- source: skills/contract-risk-review/reference/liability-cap.md
  dest: skills/contract-risk-review/reference/liability-cap.md
- source: skills/contract-risk-review/reference/penalty.md
  dest: skills/contract-risk-review/reference/penalty.md
- source: skills/contract-risk-review/reference/contract-type-guide.md
  dest: skills/contract-risk-review/reference/contract-type-guide.md
- source: skills/contract-risk-review/reference/acceptance-criteria.md
  dest: skills/contract-risk-review/reference/acceptance-criteria.md
- source: skills/contract-risk-review/reference/jurisdiction.md
  dest: skills/contract-risk-review/reference/jurisdiction.md
- source: skills/contract-risk-review/reference/auto-renewal.md
  dest: skills/contract-risk-review/reference/auto-renewal.md
- source: skills/contract-risk-review/reference/unilateral-modification.md
  dest: skills/contract-risk-review/reference/unilateral-modification.md
- source: skills/contract-risk-review/reference/ip-ownership.md
  dest: skills/contract-risk-review/reference/ip-ownership.md
- source: skills/contract-risk-review/reference/payment-schedule.md
  dest: skills/contract-risk-review/reference/payment-schedule.md
- source: skills/legal-case-search-analyzer/SKILL.md
  dest: skills/legal-case-search-analyzer/SKILL.md
- source: skills/legal-case-search-analyzer/references/report-template.md
  dest: skills/legal-case-search-analyzer/references/report-template.md
- source: skills/legal-case-search-analyzer/references/field-definitions.md
  dest: skills/legal-case-search-analyzer/references/field-definitions.md
solution_files:
- source: solution/rubrics.md
  dest: solution/rubrics.md
rubric_file_paths:
- /workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx
rubrics:
- id: '01'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx</file>中，包含6份委托书的以下内容：风险等级判定、风险项、风险说明、调整意见
  weight: 1
- id: '02'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书1的风险等级判定为中风险，涉及的风险项为：受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '03'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书2的风险等级判定为中风险，涉及的风险项为：特别授权范围过宽、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '04'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书3的风险等级判定为中风险，涉及的风险项为：受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '05'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书4的风险等级判定为高风险，涉及的风险项为：委托人身份证明缺失、受托人身份证明缺失、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '06'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书5的风险等级判定为高风险，涉及的风险项为：授权范围完全未明确、受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '07'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书6的风险等级判定为中风险，涉及的风险项为：特别授权范围过宽、转委托约定缺失、法律责任承担约定缺失
  weight: 1
- id: '08'
  content: 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx</file>中，所有委托书不包含'签字/盖章完全缺失'这一风险项
  weight: 1
---

## Prompt

根据`/workspace/input/授权委托书风控校验规则.md`，对`/workspace/input/委托书/`目录下的所有委托书进行风险情况逐项核查，输出校验报告：`/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx`。

报告中需对每份委托书进行风险等级判定，识别风险项（以上授权委托书均为起草暂未签署的版本，可排除“无签名”这一风险），给出风险说明，以及调整意见。

请直接执行，不要向我提问或要求决策，遇到问题自行处理。

## Grading Criteria

- [ ] [01] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx</file>中，包含6份委托书的以下内容：风险等级判定、风险项、风险说明、调整意见
- [ ] [02] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书1的风险等级判定为中风险，涉及的风险项为：受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
- [ ] [03] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书2的风险等级判定为中风险，涉及的风险项为：特别授权范围过宽、转委托约定缺失、法律责任承担约定缺失
- [ ] [04] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书3的风险等级判定为中风险，涉及的风险项为：受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
- [ ] [05] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书4的风险等级判定为高风险，涉及的风险项为：委托人身份证明缺失、受托人身份证明缺失、转委托约定缺失、法律责任承担约定缺失
- [ ] [06] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书5的风险等级判定为高风险，涉及的风险项为：授权范围完全未明确、受托人无相应资质、转委托约定缺失、法律责任承担约定缺失
- [ ] [07] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx中，委托书6的风险等级判定为中风险，涉及的风险项为：特别授权范围过宽、转委托约定缺失、法律责任承担约定缺失
- [ ] [08] 输出的<file>/workspace/input/委托书/委托书风控校验/授权委托书风险提示报告.docx</file>中，所有委托书不包含'签字/盖章完全缺失'这一风险项
