# officebench-2-20-final-perfect v0.3 adaptation record

Status: `harbor_replayed`; difficulty remains `unassigned`; `calibration_status=uncalibrated`.

## Upstream seed, evaluator, and license

- Dataset: OfficeBench task `2-20/2`, Git commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- Frozen Query/Evaluator arguments: `adaptation/source-task.json`, SHA256 `aaad40b592a9cf8f9a940276c03b0e3c7de8d9e55642535e51a2df5d2c9e68c9`.
- Frozen input: upstream `tasks/2-20/testbed/data/CS161_Class_Roster.xlsx`, SHA256 `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7`. The upstream, `data/input_files`, and `environment/input_files` copies are byte-identical.
- Evaluator implementation consulted: upstream `utils/evaluate.py`, SHA256 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`; the relevant containment semantics are frozen in `adaptation/upstream-evaluator.py` for audit only.
- License: Apache-2.0. `adaptation/LICENSE-OfficeBench` is byte-identical to the upstream `LICENSE`, SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`.

## Full context read before authoring

Every existing package file was read directly: instruction, TOML, rubric, Dockerfile, both input copies, manifests, source task/license, adaptation files, metadata, solver/wrapper, Judge/test wrapper/expectations, and both old DOCX outputs. The old Golden's OOXML and its LibreOffice page were inspected. The workbook was read cell-by-cell with styles and properties: `Sheet1`, 10 rows × 10 columns, no formulas, merged cells, hidden dimensions, defined names, freeze panes, or filters. The old Golden had two Normal-style paragraphs, no tables, links, bookmarks, equations, headers, footers, or traceability.

The exact upstream `2-20/2` task JSON, related `2-20` subtasks, source workbook, OfficeBench README, license, Git revision, and full `evaluate_contain` implementation were also inspected. Upstream scoring checked six positive keywords only and could not detect a 97-point false inclusion, missing population rows, source-cell errors, wrong calculations, false boundaries, flattened pages, or broken Word-native objects.

## Input truth ledger

The exact predicate is numeric `Final = 100`; two of nine rows satisfy it:

| Source row | Student | Standing | Major | HW1 | HW2 | Midterm | HW3 | Final | Pre-Final avg | Overall | Range | Final rank | Lift |
|------------|---------|----------|-------|----:|----:|--------:|----:|------:|--------------:|--------:|------:|------------|-----:|
| `Sheet1!6` | `S1004 Fufowot Dipuj` | Senior | ECE | 9 | 94 | 83 | 54 | 100 | 60.00 | 68.00 | 91 | 1 of 9 (tie) | +40.00 |
| `Sheet1!9` | `S1007 Jehecat Cani` | Freshman | Math | 46 | 74 | 36 | 81 | 100 | 59.25 | 67.40 | 64 | 1 of 9 (tie) | +40.75 |

Final values in row order are `1, 19, 66, 19, 100, 2, 97, 100, 81`: mean `53.9`, median `66.0`, perfect `2`, below `7`, above `0`, missing/non-numeric `0`, perfect rate `22.2%`, next-highest `97`. Distribution `0-19 / 20-39 / 40-59 / 60-79 / 80-99 / 100` is `4 / 0 / 0 / 1 / 2 / 2`. Nearest excluded rows are S1006 (`97`, gap 3), S1008 (`81`, gap 19), and S1002 (`66`, gap 34).

## Adaptation boundary and sibling differentiation

Version 0.3 supersedes the two-line keyword list in v0.2. Apache-2.0 permits the derivative adaptation. Retained source contributions are the workbook facts, exact Final=100 intent, Query, and evaluator record. CoworkEvoBench contributes the exception-verification workflow, all source locators and calculations, native Word objects, population/boundary/distribution controls, input-driven Oracle, partial-credit rubric, and directed failures.

The shared-workbook siblings are intentionally different. `officebench-2-20-all-cs161` is a full-class identity/control book; `officebench-2-20-cs-major-list` is a privacy-limited CS advising queue with comments; this task is an exact-threshold exception packet with one certificate per qualifying row, native equations, nearest-excluded analysis, and decision closure. The related `3-60` folder task must use a separate Word capability target. All remain in leakage group `officebench-cs161-roster`.

## Golden, Oracle, and scoring boundary

The only deliverable is `perfect_final_verification_packet.docx`. The dedicated solver constructs editable text, twelve native tables, two OMML equations, two bookmarks/internal links, a page header, automatic page numbers, repeating headers, and non-splitting rows. The Judge reads neither Golden nor solver markers: it independently rebuilds selection, source order, values, locators, derived facts, threshold boundary, distribution, and closure from the frozen XLSX. It also parses OOXML safety/native structures and renders through LibreOffice.

Golden and fresh replay must repeatedly score 1.0. Empty, corrupt, flattened, extra-output, threshold, certificate, evidence, calculation, aggregation, boundary, equation, and navigation failures must fail. A deliberately incomplete closure sample must receive a stable nonzero partial score. Oracle success does not establish real-Agent difficulty; difficulty remains unassigned pending separate multi-Harness calibration.

## Visual QA, local validation, and Harbor replay

LibreOffice 7.3.7.2 renders Golden and fresh replay as four portrait Letter pages: population ledger, two separately paged certificates, and boundary/distribution/closure. Golden/fresh, a 97-point false inclusion, a wrong pre-Final average, the medium incomplete closure, and the flattened-image case account for 22 rendered pages. Every page was checked at original render resolution; unchanged negative pages were additionally proven byte-identical to their Golden page renders. Valid pages have no blank page, clipping, overlap, off-page content, broken table flow, missing font, or unreadable text. The intentionally flattened case is visibly a title page followed by a black raster page.

Golden and fresh are byte-identical, SHA256 `98474d2f96fa02ac05c768cf60569d689afade1d412a12f1671f3dbd75dbe2b9`. Five Golden and five fresh Judge runs each score `1.0`; all ten JSON payloads are byte-identical, SHA256 `62d3fa63a58a278012b40ab4683b535a814716c5d3742c105044316e6049d4ea`. LibreOffice DOCX round-trips of Golden and fresh each score `1.0`; the Judge accepts LibreOffice's semantic minus-sign normalization only while still requiring native OMML equations.

Directed calibration rewards are: empty `0.0`, corrupt `0.0`, flattened `0.0`, unexpected extra DOCX `0.0`, 97-point false inclusion `0.35`, missing ledger row `0.35`, wrong certificate subject `0.45`, wrong certificate score/locator `0.65`, wrong average or missing equations `0.76`, wrong mean/distribution/boundary/internal links `0.78`, and medium missing one Closure control `0.78`. Every non-Golden case fails the `0.82` pass threshold, while the medium case remains strictly between zero and Golden.

The host's shared BuildKit queue stalled before executing this task, and the Debian default mirror also stalled during a diagnostic build. The task Dockerfile now uses the Aliyun HTTPS Debian mirror with bounded retries/timeouts. The environment was prebuilt through the classic Docker backend with host networking; both official Harbor 0.21.0 jobs then ran with Docker, `force_build=true`, `n_concurrent_trials=1`, offline Oracle, the full verifier, and complete artifacts. A job `docx-hardening-b8-perfect-final-a-20260908-pass` / trial `officebench-2-20-final-perfect__u8UEkiq` and B job `docx-hardening-b8-perfect-final-b-20260908-pass` / trial `officebench-2-20-final-perfect__SQPVciL` each scored `1.0` with zero exceptions. Both report task checksum `adf165fa9da50a9dec16cd17c1d8b97e5f407befabc45bfc43fba0514349d30d` and digest `sha256:f119259a0274c94ef14dff13728ae189c896d16fd01b992d107cd42b062b3a0b`; their DOCX, Judge JSON, and artifact manifest are pairwise byte-identical.
