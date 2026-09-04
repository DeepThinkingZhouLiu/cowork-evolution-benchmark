# sb2-debugging-04-03-model-structure-repair adaptation record

Status: `pending_human_review`; author hypothesis `L4` / `expert`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Debugging/04_03`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `e141b511b8e2da0a2709f44f1133d476846cb34878beab4998e7c7c638be3fb9` and upstream Golden SHA256 is `e4678dfcde8801143ba0262a2877c14667c3b12bfc64843f3bf912df0f26adb3`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

保留上游结构恢复与历史事实，补充输出路径、单位证据和保留要求；未增加业务事实。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 413 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

四处结构恢复必须先于跨七张下游工作表的公式修复，错误具有强顺序性和大范围级联；预计约 90 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
