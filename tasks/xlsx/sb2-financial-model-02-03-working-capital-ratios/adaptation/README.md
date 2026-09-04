# sb2-financial-model-02-03-working-capital-ratios adaptation record

Status: `pending_human_review`; author hypothesis `L3` / `advanced`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Financial_Model/02_03`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `b074b65f57ff82918a5ac56b79a39ebfc3ca937ddf674b6834642b1fc2fc2634` and upstream Golden SHA256 is `d2b71c081d3fb124ccf41ebe357f7b92439bab93b3e0cb77279e47289ca3862f`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

忠实翻译并明确唯一输出；未改变上游目标。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 35 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

三个工作表、六个指标族和 35 个公式目标形成跨报表依赖链；预计熟练分析师约 45 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
