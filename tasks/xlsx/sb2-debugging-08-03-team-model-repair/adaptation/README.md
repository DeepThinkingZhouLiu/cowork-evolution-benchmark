# sb2-debugging-08-03-team-model-repair adaptation record

Status: `pending_human_review`; author hypothesis `L4` / `expert`; `calibration_status=uncalibrated`.

## Selection and source

Selected from `icyCreater/SpreadsheetBench-2-sample` at immutable revision `d8a14da1898e48d07c8a9ff8974d0effc9302ece`, task `Debugging/08_03`. The dataset is CC-BY-SA-4.0; the license text is retained. Input SHA256 is `91967c0fe1ff863be8496b4f15038554d2119ebc6b2a58a116eb77b8dad2535f` and upstream Golden SHA256 is `c1e64ce307674d7db322b439e0f3d4c8b82586b5856423cc0f7c1ccfd47a6a4f`. The input, original task record, Golden, dataset manifest and closest local adaptations were inspected.

## Adaptation decisions

把上游泛化审计请求收敛到从输入与明确匹配规则可验证的结构和公式链，不评分命名偏好等主观项。 The task remains a single-workbook workflow and now has exactly one native XLSX output. The Judge has 317 scored target cells split into three business groups. It scores group completion proportionally, so one mismatch does not zero the whole workbook, and it does not read the upstream or local Golden.

## Difficulty hypothesis

两处删除行导致收入、运营模型和三套 DCF 估值链级联，且需跨表头匹配与可维护公式；预计约 90 分钟。 The tier is an author hypothesis, not a rollout result. Expected failure modes are omitted formula families, broken cross-sheet references, hardcoded substitutes and collateral edits. A pass@1 outside the target tier range, after Judge/environment audit, falsifies this label.

## Validation and human review packet

Run static package and asset validation, fresh replay, checked-in Golden, empty, corrupt, unchanged-input, medium-defect and targeted-defect tests. LibreOffice recalculation and relevant-sheet inspection remain required. Reviewers should inspect the instruction diff, evidence ledger, target groups, formula equivalence policy, preserved workbook features and CC-BY-SA attribution. No reviewer decision, calibrated tier or release status is claimed here.

Known risk: Excel-specific array/data-table/chart behavior may not fully replay in LibreOffice. The deterministic Judge therefore scores the requested formula/structure semantics and preservation separately from optional visual inspection.
