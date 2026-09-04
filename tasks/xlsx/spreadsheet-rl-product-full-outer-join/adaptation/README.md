# spreadsheet-rl-product-full-outer-join adaptation record

Status: `pending_human_review`; author hypothesis `L3` / `advanced`; `calibration_status=uncalibrated`.

## Evidence and source

The task was selected from Spreadsheet-RL, immutable revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, record `excelforum/excel-general/task-job-68594670fd5aafc2`. The upstream instruction is retained in `adaptation/source-instruction.json`; the input SHA256 is `5d6fad7006decb19d9eb960c5bbea931dfb5595f0d69d2a873e293648f838fb5` and the audit-only public target SHA256 is `ffa66fd34b203968ebafdf6f86a1c2dd99b27b7da4bc88c4d25871d5b1377f63`. Source license evidence is retained in `adaptation/LICENSE-CC-BY-SA-4.0.txt`.

The visible input was inspected field-by-field. The source task asks for one spreadsheet workflow, and this adaptation keeps the same input semantics while fixing the output path and enforcing one native XLSX artifact. No unrelated business facts were added.

## Adaptation and Judge boundary

The checked-in Golden is retained only for audit and reference-solution comparison. `tests/judge.py` independently derives categories, transformed rows, hour totals, join keys, ranking values, and preservation checks from `data/input_files`. It does not load `solution/` or compare workbook hashes. Negative cases required before release are empty output, corrupt XLSX, unchanged input, one changed core cell, and one missing formula/object.

## Difficulty hypothesis

50 minutes of competent work is expected. The hypothesis is based on a two-key full outer join with two-sided measures and coalesced fields. This is an author hypothesis only; rollout may falsify it. Expected failure modes are omitted rows, hardcoded values, wrong tie/duplicate handling, formula incompatibility and collateral edits.

## Human review packet

Review the source license and attribution, input/target previews, original-to-adapted instruction, evidence ledger, Golden replay, empty/corrupt/unchanged/medium-defect results, and the LibreOffice recalculation/rendering result. Keep this task `pending_human_review` until a human approves the task and calibrates the tier.
