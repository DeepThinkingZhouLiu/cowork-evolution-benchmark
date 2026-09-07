# Controlled OCR package adaptation record

Status: `harbor_replayed`. Difficulty remains `unassigned`; no real-agent baseline is inferred from Oracle success.

## Frozen upstream

- Dataset: OfficeBench task `3-22/0` at Git commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`.
- License: Apache-2.0, copied byte-for-byte to `LICENSE-OfficeBench`.
- Frozen original query: `Scan PDF homework and save as homework.docx for editing, save homework file into folder named homework`.
- Frozen original evaluator: two `evaluate_file_exist` calls for `data/homework.docx` and `data/homework/homework.docx`; it contains no content, structure, or editability checks.
- Runtime source: the upstream `tasks/3-22/testbed/data/homework.jpg`, 1700 x 2200 pixels, SHA256 `1185885b07d4d5e17e23393695fe17775cb829aec4adeea366327dd4bf7953bf`.

The original evaluator explains the two-path requirement but is not adequate as a correctness oracle. Version 0.3 keeps both DOCX paths and replaces the old byte-identical, single-paragraph OCR copies with a controlled master plus a role-specific QA working copy.

## Capability redesign and supersession

This task supersedes `officebench-3-22-homework-folder@0.2.0`. The two outputs now form one governed Word package:

- `homework.docx` is a four-page controlled digitization master. It includes the canonical source payload, native equation and pipeline table, an 18-row block register, native-object register, cross-copy contract, source-exception ledger, and master closure controls.
- `homework/homework.docx` is a distinct four-page QA working copy. It carries the same canonical payload, six source-bound ambiguity issues with native Word comments, and four unlocked rich-text content controls for status, reviewer, disposition, and release gate.
- Both copies include role-specific headers, footers, page fields, bookmarks, internal links, custom properties, and a shared canonical payload digest. Their source payload must match while their governance layers and bytes must differ.
- The absent Question 2 diagram is disclosed rather than invented. The image itself is not embedded, so both outputs remain editable and do not use a flattened page as a substitute.

The scan is shared with `officebench-2-6-homework-ocr`, so both tasks use leakage group `officebench-homework-ocr-shared`. The capability targets are deliberately different: `2-6` evaluates a single faithful reconstruction with evidence/fingerprint audit; `3-22` evaluates synchronized multi-copy governance, Word comments, content controls, and role separation.

## Truth construction

`package-truth.json` is a task-specific ledger transcribed from the frozen scan and independently copied into `solution/` and `tests/`. Its SHA256 is `271e323474d9532977d4f392bef3012daf9b48e7de2eab42da6263d59b51113e`. The canonical payload digest, recomputed from its header, instructions, section heading, questions, and pipeline, is `83425e9072d022c0a797826c1d6ebf38884fa59ffb4a638c88f8cd9903cf9848`.

The Judge reads the runtime JPEG, checks its bytes and dimensions, loads the frozen truth, and recomputes every expected entity. It never opens Golden, compares Golden hashes, reads a solution marker, or accepts candidate self-reporting. The eight criteria separately cover package validity, exact ordered source content, pipeline/equation fidelity, copy synchronization and properties, master controls, review issues/comments, content controls, and native structure/rendering. Business-critical failures apply documented caps; partial entity and cell scores remain possible.

## Local validation evidence

- Golden master SHA256: `b6ef6ea83066f97142d5a2dadac155bf94816291d2ce36518d39290f5a4c9df0`.
- Golden review SHA256: `fd2006d51dba1158d6cdd5307288902e246af92fd96ea1110a2aa3ec576907b2`.
- Five Golden evaluations and five independent fresh builds each scored `1.000000`; fresh outputs were byte-identical to their respective Golden copies and every final result JSON had the same hash.
- Twenty-three task-specific negative artifacts all failed. Empty, corrupt, flattened, missing-copy, and extra-file cases scored `0`; factual, synchronization, native-object, role, comment, control, navigation, and disclosure failures scored from `0.32` to `0.996`; the deliberately incomplete review-queue sample scored `0.78`.
- Fixed LibreOffice rendered both Golden copies as exactly four portrait, nonblank, correctly ordered pages. All eight pages were inspected at original resolution: no blank page, clipping, overflow, broken table continuation, unreadable font, or footer loss was observed.
- LibreOffice open/save/export roundtrip retained both four-page layouts, source content, comments, properties, tables, bookmarks, links, headers, and footers. It removed the four `w:tag` values while retaining control aliases, values, and unlocked state, so the strict Judge correctly scored the roundtrip `0.82` instead of weakening the authored requirement.
- Selected causal negatives were rendered and visually inspected, including payload deletion, pipeline corruption, unsynchronized identity, fabricated diagram claim, missing control, missing role header, incomplete review queue, wrong issue disposition, and a two-page flattened substitute.

## Harbor replay and remaining gate

Official Harbor Docker Oracle A/B replay ran with `n-concurrent=1`. Jobs `docx-hardening-b8-controlled-ocr-package-a-20260908` and `docx-hardening-b8-controlled-ocr-package-b-20260908` each completed one trial at reward `1.0` with zero exceptions. Both locks used task digest `sha256:cc5f80a2b03981aa8d667cd597349145bf2a48d046bd8edb8fad6b2c57f83355`; both manifests and Judge payloads were identical, and both artifact pairs matched Golden byte-for-byte. These replays prove reproducibility of the authored solution only; difficulty remains unassigned until later real-agent calibration.
