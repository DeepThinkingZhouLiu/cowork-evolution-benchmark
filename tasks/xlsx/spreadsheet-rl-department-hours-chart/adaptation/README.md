# spreadsheet-rl-department-hours-chart adaptation record

Status: `pending_human_review`; author hypothesis `L2` / `intermediate`; `calibration_status=uncalibrated`.

## Evidence and source

The task was selected from Spreadsheet-RL, immutable revision `6f274cd4cbc27e656d7ddae7c3b591669866dc84`, record `excelforum/excel-general/task-job-1813a28727a048cb`. The upstream instruction is retained in `adaptation/source-instruction.json`; the input SHA256 is `79b9fd32457920f7607ac721bf245013c05f07ffc97002af3a7c78b6373526ad` and the audit-only public target SHA256 is `f4205998c41aa7bf8d04740b34db1a1576bec33aa2004e87b54d82b34244ad09`. Source license evidence is retained in `adaptation/LICENSE-CC-BY-SA-4.0.txt`.

The visible input was inspected field-by-field. The source task asks for one spreadsheet workflow, and this adaptation keeps the same input semantics while fixing the output path and enforcing one native XLSX artifact. No unrelated business facts were added.

## Adaptation and Judge boundary

The checked-in Golden is retained only for audit and reference-solution comparison. `tests/judge.py` independently derives categories, transformed rows, hour totals, join keys, ranking values, and preservation checks from `data/input_files`. It does not load `solution/` or compare workbook hashes. Negative cases required before release are empty output, corrupt XLSX, unchanged input, one changed core cell, and one missing formula/object.

## Difficulty hypothesis

55 minutes of competent work is expected. The hypothesis is based on mixed time encodings, 22 departments, 28 months and a multi-series chart. This is an author hypothesis only; rollout may falsify it. Expected failure modes are omitted rows, hardcoded values, wrong tie/duplicate handling, formula incompatibility and collateral edits.

## Human review packet

Review the source license and attribution, input/target previews, original-to-adapted instruction, evidence ledger, Golden replay, empty/corrupt/unchanged/medium-defect results, and the LibreOffice recalculation/rendering result. Keep this task `pending_human_review` until a human approves the task and calibrates the tier.
