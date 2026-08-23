# 公开种子题目录

> 本文件由 `tools/task_generation/build_catalog.py` 生成。当前内容是公开开发种子题，不等同于论文密封测试集。

共 42 题，覆盖 42 个任务家族；包含 95 个输入文件和 59 个 Golden 交付物。

| 领域           | 任务家族                         | 任务编号                                            | 难度           | 输入/输出 | 能力标签                          |
| ------------ | ---------------------------- | ----------------------------------------------- | ------------ | ----- | ----------------------------- |
| cross_office | `extract_analyze_present`    | `cross-001-extract-analyze-present`             | expert       | 2/2   | K1, K3, K4, K5, K8, K10, K13  |
| cross_office | `research_to_deliverables`   | `cross-002-research-to-deliverables`            | advanced     | 3/3   | K1, K2, K5, K6, K8, K10, K13  |
| cross_office | `shared_source_multi_output` | `cross-003-shared-source-multi-output`          | advanced     | 1/3   | K1, K6, K8, K9, K11, K13      |
| cross_office | `update_propagation`         | `cross-004-update-propagation`                  | expert       | 4/3   | K7, K10, K11, K13             |
| cross_office | `cross_file_consistency`     | `cross-005-cross-file-consistency`              | expert       | 4/2   | K1, K2, K5, K10, K12          |
| cross_office | `template_driven_package`    | `cross-006-template-driven-package`             | expert       | 4/3   | K4, K7, K9, K11, K13          |
| cross_office | `review_and_revision`        | `cross-007-review-and-revision`                 | expert       | 3/2   | K2, K5, K7, K8, K10, K12, K13 |
| cross_office | `workspace_orchestration`    | `cross-008-workspace-orchestration`             | expert       | 4/3   | K1, K2, K4, K5, K10, K11, K13 |
| docx         | `material_to_report`         | `docx-001-material-to-report`                   | advanced     | 3/1   | K1, K2, K5, K6, K10           |
| docx         | `outline_to_document`        | `docx-002-outline-to-document`                  | intermediate | 2/1   | K1, K2, K6, K9                |
| docx         | `template_filling`           | `docx-003-template-filling`                     | advanced     | 3/1   | K1, K7, K9, K10               |
| docx         | `editing_rewriting`          | `docx-004-editing-rewriting`                    | intermediate | 2/1   | K2, K6, K7, K9                |
| docx         | `redlining_review`           | `docx-005-redlining-review`                     | expert       | 2/1   | K2, K7, K9, K12               |
| docx         | `formatting_fidelity`        | `docx-006-formatting-fidelity`                  | intermediate | 2/1   | K7, K8, K9, K12               |
| docx         | `merge_compare`              | `docx-007-merge-compare`                        | advanced     | 2/2   | K2, K7, K10, K11, K12         |
| docx         | `document_extraction`        | `docx-008-document-extraction`                  | advanced     | 3/1   | K1, K2, K5, K6, K10           |
| ppt          | `material_to_deck`           | `ppt-001-community-canteen-retrospective`       | advanced     | 3/1   | K1, K2, K4, K7, K9            |
| ppt          | `outline_to_deck`            | `ppt-002-nebula-launch-outline`                 | intermediate | 3/1   | K1, K3, K7, K9                |
| ppt          | `data_to_deck`               | `ppt-003-east-china-retail-weekly`              | advanced     | 2/1   | K2, K4, K5, K7, K9            |
| ppt          | `template_based_generation`  | `ppt-004-customer-success-template`             | advanced     | 3/1   | K2, K4, K6, K7, K10           |
| ppt          | `content_editing`            | `ppt-005-renewal-review-revision`               | advanced     | 2/1   | K2, K6, K8, K10               |
| ppt          | `visual_editing`             | `ppt-006-warehouse-safety-layout`               | intermediate | 2/1   | K6, K7, K8, K10               |
| ppt          | `deck_polishing`             | `ppt-007-quarterly-ops-polish`                  | advanced     | 2/1   | K4, K6, K7, K8, K10           |
| ppt          | `deck_consolidation`         | `ppt-008-product-lines-consolidation`           | expert       | 3/1   | K2, K4, K6, K8, K10, K13      |
| ppt          | `audience_adaptation`        | `ppt-009-architecture-executive-adaptation`     | expert       | 2/1   | K1, K3, K6, K8, K10           |
| ppt          | `localization`               | `ppt-010-apac-launch-localization`              | advanced     | 2/1   | K3, K6, K8, K10               |
| ppt          | `presentation_package`       | `ppt-011-volunteer-conference-package`          | expert       | 4/3   | K1, K2, K4, K7, K9, K12       |
| ppt          | `deck_audit`                 | `ppt-012-board-deck-audit`                      | expert       | 3/1   | K2, K4, K6, K8, K11, K13      |
| xlsx         | `analysis_to_report`         | `xlsx-analysis-to-report-service-review`        | advanced     | 1/2   | K1, K2, K5, K6, K8, K13       |
| xlsx         | `charting`                   | `xlsx-charting-growth-funnel`                   | intermediate | 1/1   | K2, K5, K8, K9                |
| xlsx         | `dashboard`                  | `xlsx-dashboard-sales-operations`               | advanced     | 1/1   | K2, K4, K5, K8                |
| xlsx         | `data_cleaning`              | `xlsx-data-cleaning-crm-leads`                  | intermediate | 2/1   | K2, K3, K10, K12              |
| xlsx         | `document_to_table`          | `xlsx-document-to-table-procurement-quotes`     | intermediate | 1/1   | K1, K2, K3, K6                |
| xlsx         | `formula_authoring`          | `xlsx-formula-authoring-order-margin`           | intermediate | 1/1   | K2, K4, K7                    |
| xlsx         | `model_building`             | `xlsx-model-building-saas-plan`                 | advanced     | 1/1   | K2, K4, K5, K8                |
| xlsx         | `multi_workbook_integration` | `xlsx-multi-workbook-integration-branch-orders` | advanced     | 4/1   | K2, K3, K10, K12              |
| xlsx         | `pivot_analysis`             | `xlsx-pivot-analysis-retail-channel`            | intermediate | 1/1   | K2, K3, K5, K8                |
| xlsx         | `scenario_analysis`          | `xlsx-scenario-analysis-logistics-pricing`      | advanced     | 1/1   | K2, K4, K5, K8                |
| xlsx         | `spreadsheet_audit`          | `xlsx-spreadsheet-audit-project-costs`          | advanced     | 2/1   | K2, K5, K10, K12              |
| xlsx         | `spreadsheet_repair`         | `xlsx-spreadsheet-repair-budget-errors`         | advanced     | 1/1   | K2, K4, K7, K12               |
| xlsx         | `table_transformation`       | `xlsx-table-transformation-regional-sales`      | intermediate | 1/1   | K2, K3, K4                    |
| xlsx         | `template_completion`        | `xlsx-template-completion-headcount-plan`       | intermediate | 2/1   | K1, K4, K7, K9, K10           |
