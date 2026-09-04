# sb2-financial-model-20-03-operations-rollforward adaptation record

Status: `pending_human_review`; author hypothesis `L3` / `advanced`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Financial_Model/20_03`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `eaf48f79443270b2330075c1489cf47ab51bebe25b22ffcb70392966f422d0d2` and upstream Golden SHA256 is `c6daa76101dbaa63f90615aadb9e171a8e841da2ec61252472bfaccf152c84d2`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

输入中的 H:AW 月度 Working Capital 已完成且与 Golden 一致，适配题面明确只补年度 C:F，保留月度滚动。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 18 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

四个工作表和四类运营模型公式形成年度/月度联动；预计约 45 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
