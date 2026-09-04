# officebench-1-13-cs161-prerequisites

Status: `pending_human_review`; initial difficulty: `basic` (L1); `calibration_status=uncalibrated`.

## Upstream evidence

- OfficeBench task `1-13/0` at immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original Query is copied verbatim to `adaptation/source-task.json`: `Given an Excel containing student's previous courses, does he/she satisfy the prerequiste of CS161 if prerequiste classes are CS101, CS 102, Math 151, and Math 152? Answer should be Yes or No.`
- The source evaluator expected a TXT answer. The minimal format adaptation changes only the answer container to an editable DOCX; the question, source workbook, and answer semantics are unchanged.
- Input SHA256 is recorded in `metadata/input-manifest.json`; the same bytes are copied under `environment/input_files/`.

## Validation boundary

The Golden and reference solver are generated from the visible workbook. The Judge independently recomputes the target and checks OOXML readability, exact answer structure, no placeholders, and input hash. This package remains a candidate until human review of the source preview, Golden rendering, and negative cases.

## Author-side validation (2026-09-04)

- `validate-task` and `validate-task-assets`: pass.
- Checked-in Golden and a fresh replay from packaged inputs: reward `1.000000`; each was scored five times with identical deterministic results.
- Missing output and corrupt bytes: fail, with reward between `0.000000` and `0.100000`. A screenshot-only DOCX also fails.
- A non-empty targeted defect that removes one required fact, paragraph, or table row: reward `0.43` and fail. For source-edit tasks, the unchanged input also fails.
- LibreOffice `7.3.7.2` rendered Golden and fresh replay to 1 page each. Text and pixel-boundary checks found no unintended blank page or clipping, and the rendered contact sheet was inspected.

These are author-side checks, not human acceptance or rollout calibration. Status remains `pending_human_review`; `calibration_status=uncalibrated`.
