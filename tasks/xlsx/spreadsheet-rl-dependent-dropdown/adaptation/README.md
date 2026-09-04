# spreadsheet-rl-dependent-dropdown adaptation record

Status: `pending_human_review`; author hypothesis `L2` / `intermediate`; `calibration_status=uncalibrated`.

## Evidence and source

The task was selected from Spreadsheet-RL, immutable revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, record `excelforum/excel-general/task-job-2d603b0f50935350`. The upstream instruction is retained in `adaptation/source-instruction.json`; the input SHA256 is `4bd7e6e036fe06ccabba76dbdb5cdc76d117f4ba4b36ca4b7b95b549baa76ea2` and the audit-only public target SHA256 is `dd2726369e21e4010138d4deeade550eff93462c5c913621c4ef4d117855bcc2`. Source license evidence is retained in `adaptation/LICENSE-CC-BY-SA-4.0.txt`.

The visible input was inspected field-by-field. The source task asks for one spreadsheet workflow, and this adaptation keeps the same input semantics while fixing the output path and enforcing one native XLSX artifact. No unrelated business facts were added.

## Adaptation and Judge boundary

The checked-in Golden is retained only for audit and reference-solution comparison. `tests/judge.py` independently derives categories, transformed rows, hour totals, join keys, ranking values, and preservation checks from `data/input_files`. It does not load `solution/` or compare workbook hashes. Negative cases required before release are empty output, corrupt XLSX, unchanged input, one changed core cell, and one missing formula/object.

## Difficulty hypothesis

35 minutes of competent work is expected. The hypothesis is based on multiple validation objects, named ranges and dependent formulas. This is an author hypothesis only; rollout may falsify it. Expected failure modes are omitted rows, hardcoded values, wrong tie/duplicate handling, formula incompatibility and collateral edits.

## Human review packet

Review the source license and attribution, input/target previews, original-to-adapted instruction, evidence ledger, Golden replay, empty/corrupt/unchanged/medium-defect results, and the LibreOffice recalculation/rendering result. Keep this task `pending_human_review` until a human approves the task and calibrates the tier.
