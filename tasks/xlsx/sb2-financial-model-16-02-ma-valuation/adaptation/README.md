# sb2-financial-model-16-02-ma-valuation adaptation record

Status: `pending_human_review`; author hypothesis `L4` / `expert`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Financial_Model/16_02`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `4567a1dab48e5b6876f295fde3a82ff4d12bbe75aa25475c7e5702eb6569a867` and upstream Golden SHA256 is `42cbd23372cc480fd2184e61930f7bfe1eaede4cc6f069ea3de0e0e82f8d20f5`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

忠实翻译并明确唯一输出；不要求修复模型中与任务无关的既存引用。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 30 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

估值、IRR、终值、协同成本和两张财务报表之间存在长链专业依赖，30 个目标跨三个模型区域；预计约 75 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
