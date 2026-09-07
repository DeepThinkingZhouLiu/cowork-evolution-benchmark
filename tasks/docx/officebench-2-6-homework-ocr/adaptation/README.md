# officebench-2-6-homework-ocr v0.3 adaptation record

Status: `harbor_replayed`; Harbor Oracle A/B passed. Difficulty remains `unassigned`; no author-side difficulty tier is frozen.

## Frozen upstream evidence

- Dataset: OfficeBench task `2-6/0`, Git revision `b978b808667c32b52ce19a67ce1def1de9ae02b7`, Apache-2.0.
- Original query: `Scan PDF homework to homework.docx for editing`. The source actually supplies one JPEG scan, so v0.3 says JPEG rather than repeating the inaccurate PDF label.
- Original evaluator: one `evaluate_contain` call for `CSE`, `240A`, and `Homework`. It is retained verbatim in `source-task.json` but is not used as the v0.3 Oracle.
- Exact source snapshot: `homework.jpg`, 1700 × 2200 pixels, SHA256 `1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf`. Upstream, package, and runtime copies are byte-identical.
- The complete Apache-2.0 license is retained as `LICENSE-OfficeBench`. The upstream README and executable evaluator were also read in full; their hashes are in `evidence-ledger.md`.

## Why this is a remake rather than a format tweak

The v0.2 output was one noisy OCR paragraph and the Judge mostly inherited the three-keyword upstream check. Version 0.3 replaces that weak contract with three distinct editable Word pages:

1. an exact source reconstruction with five ordered instruction bullets, all three visible questions, both repeated notes, a native 7 × 16 pipeline matrix, three visible dependency arrows, and a native OMML timing equation;
2. a source anchor and an 18-row transcription fingerprint ledger; and
3. an independent pipeline audit, explicit source-conflict register, disambiguation conventions, and package closure controls.

The source says there are two questions while visibly numbering prompts 1, 2, and 3. It also refers to a second diagram that is absent from the supplied scan. Both conflicts are preserved and explained; no missing diagram, answer, or grading fact is invented. The scan is evidence only and is not embedded as a flattened output page.

## Truth boundary and evaluator independence

`transcription-truth.json` is a hand-verified ledger derived from the frozen scan. It records every visible text segment, the six pipeline rows across cycles 1–15, the three down-arrow relations, the worked timing comparison, and the two source conflicts. Equivalent evaluator snapshots live beside `solve.py` and `judge.py` because Harbor mounts `/solution` and `/tests` separately; after removal of the adaptation-only `schema_version` field, all three JSON objects are equal. The solution/test copies are byte-identical with SHA256 `f9ac0f3acea88968d77563bdf4dde0aaf860d2a88d597f4581da682d6aa702cf`.

The task-specific Judge reads the frozen input and truth snapshot, independently verifies the JPEG hash and dimensions, parses DOCX/OOXML, and checks each segment, table row/cell, evidence fingerprint, conflict treatment, native object, navigation target, custom property, header/footer, and rendered three-page contract. It rejects external relationships, embedded media, executable parts, extra DOCX deliverables, and flattened-image substitutions. It never reads Golden bytes, compares a Golden hash, or trusts candidate self-reporting.

## Leakage and family separation

This task shares the scan only with `officebench-3-22-homework-folder`, so both retain leakage group `officebench-homework-ocr-shared`. This task's capability goal is forensic single-scan reconstruction plus a pipeline/evidence audit. Task 3-22 must use a different document-workflow goal and cannot be a renamed or relocated copy. Version 0.3 supersedes the weak v0.2 OCR-paragraph contract under the same stable task ID.

## Local validation evidence — 2026-09-08

- Golden, checked-in solution, and five independent fresh replays are byte-identical, SHA256 `298605033d3a28ced8225da7062fc3ec69338874a1cdc641ec9d0512153e8268`; each Judge run scores `1.000000`.
- The final LibreOffice 7.3 export is three portrait Letter pages, PDF SHA256 `56ad54d049b8c5aedf86b2f70a5355d0a52b3c0cbeb47d8678db836eae408f29`. All three pages were inspected at original raster resolution: no blank page, clipping, overlap, table spill, unreadable text, or missing reconstruction section was found.
- A LibreOffice open/save/export round trip produced DOCX SHA256 `a771e653528179464dd89a2d1b12ef08296a720844c7ae302201676438ef52f0` and a three-page PDF SHA256 `43aedcecdca9e5d731593af02dc07c04b93f19386b5009a48ab6b8de7ee960ec`; the round-tripped DOCX scores `1.000000`. Writer rewrites the multiplication glyph and odd-page header display, but the semantic equation, all body content, tables, properties, links, pagination, and readability remain intact.
- Twenty task-specific negative cases were run twice with isolated LibreOffice profiles. Every JSON result is stable and every negative fails: empty/corrupt/flattened/fabricated/extra output `0.00`; missing bullet or wrong student ID `0.30`; M1→Ml or wrong equation relation `0.42`; missing decode stall, dependency arrow, or pipeline instruction `0.55`; wrong/missing evidence row `0.68`; conflict/convention defects and the medium-quality missing closure row `0.76`; broken navigation or missing StudentID property `0.80`; missing repeating header `0.82`.
- Selected negative render pages were inspected directly, including identity, M1/Ml, arrow, fingerprint, conflict-policy, medium-quality, and flattened-output failures. Valid structured negatives remain readable and three pages; the flattened cheat renders as two invalid pages and is rejected before semantic scoring.
- Judge rendering uses a per-run isolated LibreOffice user profile. This removes observed cross-process profile contention and makes five Golden plus five evidence-negative runs byte-for-byte stable.

## Official Harbor Oracle A/B — 2026-09-08

- A: job `docx-hardening-b8-ocr-reconstruction-a-20260908`, trial `officebench-2-6-homework-ocr__zmGmBAs`, reward `1.0`, exceptions `0`.
- B: job `docx-hardening-b8-ocr-reconstruction-b-20260908`, trial `officebench-2-6-homework-ocr__3y2kPJ5`, reward `1.0`, exceptions `0`.
- Both use task digest `sha256:e9bc27302500ad6355b37e757c708092423c752bdd16c902b97aefd409c186d1` with `n-concurrent=1` and `force_build=true`.
- Both artifacts contain the required `homework.docx`, SHA256 `298605033d3a28ced8225da7062fc3ec69338874a1cdc641ec9d0512153e8268`; artifact manifests are complete, status `ok`, and byte-identical, SHA256 `3b2200d91f247c9edd3849a8c80e83402294fbabaf72c2f7e2523134acac9c2c`.
- Both Judge payloads are byte-identical, SHA256 `c0b63a8b0a9324823a437818d2d495cb1f910732a0ea3e4ca785eb10008c5210`, with all eight criteria at `1.0`.

This confirms Oracle reproducibility only. No paid three-Harness baseline was run, so real-agent difficulty remains uncalibrated.
