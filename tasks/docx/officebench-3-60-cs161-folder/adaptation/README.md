# officebench-3-60-cs161-folder v0.3 adaptation

Status: `harbor_replayed`. Difficulty remains `unassigned`; `calibration_status=uncalibrated` until independent agent baselines exist.

## Frozen upstream

- Dataset: OfficeBench, task `3-60/0`, commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`, Apache-2.0.
- Original Query: `Find all students who are taking CS161 and put the name and student ID of students into students.docx, save in folder CS161`.
- Original evaluator only checked that `data/CS161/students.docx` existed and contained five last names: `Pivozu`, `Cixarim`, `Pegoqeju`, `Xihozun`, and `Zejo`.
- `source-task.json` is the byte-identical upstream task record; `LICENSE-OfficeBench` is the byte-identical upstream license. The input workbook in both runtime locations is byte-identical to the upstream fixture.

| Frozen item | SHA256 |
|---|---|
| OfficeBench README | `04104d99107c2f109ea57d06f75541051efd4e7b5a8b2e41083315b3d21d5ed2` |
| OfficeBench LICENSE / packaged license | `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4` |
| OfficeBench `utils/evaluate.py` | `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90` |
| Upstream `tasks/3-60/subtasks/0.json` / packaged `source-task.json` | `4c648ffe87f47018910424bd20041b3a7eec829682c238d3215a517690eb64ae` |
| Upstream/package workbook | `f45c20ec06e621a6ad3b05549df9109f439c3bee09680d37c71d5087b5a418b3` |

The Apache-2.0 source permits adaptation. This remake contributes the task-specific register design, mechanical verification rules, deterministic solver, input-recomputed Judge, Golden, negative calibration suite, and audit evidence; it does not claim the synthetic OfficeBench student records as newly authored real people.

## Read evidence and old-task finding

The author read the complete original instruction, rubric, Judge, solver, adaptation, metadata, source task, license, workbook, and all three prior reference DOCX files before redesign. The workbook has one visible sheet named `Sheet1`, one header row, 123 data rows (`2–124`), nine columns, unique IDs `S1000–S1122`, no formulas, merges, hidden sheets/rows/columns, hyperlinks, or comments. The prior three reference DOCX files were byte-identical (`70f2fb5ae362f7a71e56f95b17fe0d1fdac0e039035e93aa607fc3db8c8e70c7`): each was a 123-paragraph plain name/ID dump with no tables, navigation, reconciliation, or source-row proof. Its five-keyword evaluator could not detect omissions, duplicates, reordered rows, wrong IDs, wrong scores, or image flattening.

## v0.3 capability decision

The remake keeps the original editable DOCX-only contract at `CS161/students.docx`, but changes the business goal to a long-roster verification register:

- preserve all 123 source records and all nine source fields in workbook order, partitioned into five source-row blocks;
- independently recompute overall and standing-level assessment profiles;
- derive a 56-record mechanical queue from four disclosed rules (`ZERO`, `PERFECT`, `SPREAD80`, `SWING60`) and a three-key deterministic ordering;
- reconcile block, queue, aggregate, and package controls on an explicit closure page;
- use native editable Word tables, controlled pagination, headings, bookmarks, internal links, repeating table headers, non-splitting rows, footer fields, and seven custom properties;
- state a strict boundary: flags are source checks, not grades, consequences, misconduct findings, or advising decisions.

The same review-policy JSON is frozen byte-for-byte in `adaptation/`, `solution/`, and `tests/`. Solver and Judge consume their own copies; the Judge opens the frozen workbook and recomputes every expected row, aggregate, flag, order, digest, and ledger value. It never imports the solver, reads the Golden, compares a Golden hash, or trusts candidate-reported truth.

## Same-source-family differentiation

This task shares the OfficeBench CS161 roster family with `officebench-2-20-all-cs161`, `officebench-2-20-cs-major-list`, and `officebench-2-20-final-perfect`, so metadata retains `leakage_group_id=officebench-cs161-roster`. It is not a path-only variant:

- `2-20-all-cs161` is a nine-person full-field class control book on a different workbook.
- `2-20-cs-major-list` is a three-person privacy-limited CS-major advising docket.
- `2-20-final-perfect` is a two-person exact `Final=100` verification packet.
- `3-60` is a 123-person long-register transfer, reconciliation, exception-queue, and controlled-pagination task on its own upstream fixture.

The v0.3 package supersedes only `officebench-3-60-cs161-folder@0.2.0`; the public ID remains stable because the required `CS161/students.docx` outcome still matches the upstream task.

## Exact truth ledger

- Source rows: 123; first/last IDs: `S1000` / `S1122`.
- Standing counts: Freshman 31, Sophomore 30, Junior 32, Senior 30.
- Rule totals: `ZERO=8`, `PERFECT=7`, `SPREAD80=34`, `SWING60=30`, any flag 56.
- Queue partitions: 20, 20, 16 records.
- Block digests: B1 `070f9935303d`; B2 `88abdb4e2a65`; B3 `546878816d0f`; B4 `14c19b854d00`; B5 `953a8453b231`.
- Golden: one DOCX, SHA256 `9c96a13ef36c091861bf0d973028230c7090441c47adc7baeafe16fdb8da9471`.

## Local validation and visual audit

- `validate-task`, `validate-task-assets`, task-local `compileall`, JSON parsing, input-copy equality, policy-copy equality, Golden/solution equality, and `git diff --check` pass.
- Golden and independently generated fresh output are byte-identical. Golden/fresh each receive 1.0 in five repeated Judge runs; all ten result JSON files have SHA256 `ef1f3749776ad83f5e2a71edcbd6cc488d137618805fab0e7624df7887b73130`.
- Empty, corrupt, wrong-path, extra-DOCX, and flattened-image packages receive 0.0. Twenty-three additional task-specific directed failures are rejected; the intentionally incomplete closure sample receives 0.78. All 28 calibration packages produce byte-identical Judge JSON on a second run.
- Fixed LibreOffice renders of Golden and fresh output have exactly 11 pages. Every page was inspected at original resolution for orientation, blank pages, clipping, overlap, font loss, table continuity, headers, footers, and readability.
- Nine representative negative packages were rendered and inspected on their changed pages: missing roster row, wrong source score, wrong queue order, wrong overall mean, wrong block digest, portrait-register damage, fabricated academic decision, incomplete closure, and image flattening.
- LibreOffice 7.3 open/save/export preserves 11 pages, all source and derived business facts, tables, properties, bookmarks, links, sections, and readable layout. It rewrites the simple Word footer field representation; the strict Judge therefore gives the round-tripped DOCX 0.83 while the exported PDF still visibly shows correct `PAGE / NUMPAGES` values. This interoperability loss is recorded rather than hidden.

## Harbor Docker Oracle A/B

- Harbor `0.21.0`, Docker environment, `n-concurrent=1`, `force_build=true`, Oracle agent.
- A: job `docx-hardening-b8-long-roster-register-a-20260908`, trial `officebench-3-60-cs161-folder__Lri977a`, Reward 1.0, exception 0.
- B: job `docx-hardening-b8-long-roster-register-b-20260908`, trial `officebench-3-60-cs161-folder__j6P8dq9`, Reward 1.0, exception 0.
- Both frozen runs report `task_checksum=cd77458e7fc2a39476b73b55f7ed99ee8f51f2123e633d7ad65dbfbb5381a26e` and `task_digest=sha256:7fafdcfdcdd67beea063807c7008b9c1e74f872074e4b6c8cbd10b1658436ab9`.
- Both DOCX artifacts have SHA256 `9c96a13ef36c091861bf0d973028230c7090441c47adc7baeafe16fdb8da9471`; both Judge JSON files have SHA256 `ef1f3749776ad83f5e2a71edcbd6cc488d137618805fab0e7624df7887b73130`; both artifact manifests have SHA256 `3b2200d91f247c9edd3849a8c80e83402294fbabaf72c2f7e2523134acac9c2c`.

Oracle success proves the frozen reference workflow and evaluator contract are internally consistent. It does not establish real-agent difficulty.
