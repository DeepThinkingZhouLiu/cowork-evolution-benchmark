# officebench-1-13-cs-transfer-eligibility

Status: `pending_human_review`; initial difficulty: `basic` (L1); `calibration_status=uncalibrated`.

## Upstream evidence

- OfficeBench task `1-13/1` at immutable revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Original Query is copied verbatim to `adaptation/source-task.json`: `Given an Excel containing student's previous courses, is he/she qualified for transfering major to CS if student need to take at least 5 CS courses for transfering to CS major? Answer should be 'Yes' or 'No'.`
- The source evaluator expected a TXT answer. The minimal format adaptation changes only the answer container to an editable DOCX; the question, source workbook, and answer semantics are unchanged.
- Input SHA256 is recorded in `metadata/input-manifest.json`; the same bytes are copied under `environment/input_files/`.

## Validation boundary

The Golden and reference solver are generated from the visible workbook. The Judge independently recomputes the target and checks OOXML readability, exact answer structure, no placeholders, and input hash. This package remains a candidate until human review of the source preview, Golden rendering, and negative cases.
