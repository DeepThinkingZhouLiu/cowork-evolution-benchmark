# sb2-financial-model-18-03-project-ratios adaptation record

Status: `pending_human_review`; author hypothesis `L3` / `advanced`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Financial_Model/18_03`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `fd095b2c0011d098aac379b6f8f928f79cc5a8ee4dcbdb81c6c5a88628ddf6c4` and upstream Golden SHA256 is `2406a01d43efa4fe61c12ec342d40fa12779a01ffd1b706297be7781c6bfd039`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

审计发现 DSCR!C17 和 Cash Flow!G30:L30 在输入中已完成且与上游 Golden 相同，因此删除两项伪缺失要求并明确不得改写。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 29 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

三个工作表、29 个跨年度目标，需正确处理 NM 语义和资产负债表聚合；预计约 40 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
