# Source Audit And Evidence Composition

Read this before selecting or combining upstream material.

## Source acceptance gate

For each candidate, resolve all of the following from repository files or an authoritative dataset card:

| Question | Required evidence |
|---|---|
| What did the user actually ask? | Original query, task record, or faithful session request |
| What state was visible to the user? | Workbooks, attachments, database fixture, API snapshot, email/calendar state, or deterministic reconstruction source |
| What makes an answer correct? | Input-derived facts, business rules, state assertions, evaluator, reviewed Golden, or a combination |
| Can the task run offline and repeatably? | Immutable assets or a deterministic projection from an immutable fixture |
| Can the assets be used and redistributed? | Repository license plus per-asset terms and provenance |
| Is the task suitable for a spreadsheet deliverable? | A native XLSX is a primary artifact, not a decorative export added by the author |

Reject or defer a candidate when the user intent is only a category label, inputs are unavailable, truth is unknowable, licensing is unresolved, private data cannot be removed without changing the task, or correct completion depends on a live service that cannot be frozen faithfully.

## Evidence ledger

Record a row for every retained source and derived asset:

```text
evidence_id | upstream repo/revision/path | local path | role
license basis | original/projected/derived | transformation
SHA256 | facts or requirements supported | unresolved risk
```

Use immutable commits or dataset revisions. Record the source archive/database checksum as well as projected-file checksums when a static input is extracted from a larger fixture.

`synthetic=true` describes upstream benchmark fixtures or mock enterprise data. It does not mean the CoworkEvoBench author may invent new records. Be conservative: a real public benchmark task is not proof that its spreadsheet contains production company data.

## Adaptation modes

### Direct asset adaptation

Use when the upstream source supplies a coherent query and distributable input files. Preserve the files byte-for-byte where possible. Preserve the upstream Golden, Rubric, and verifier in `adaptation/` as evidence, then audit them independently.

Permitted changes include fixed output paths, clear input filenames, offline execution boundaries, faithful translation, explicit dates replacing relative time, clarification of an ambiguity already resolved by source evidence, and removal of unsupported upstream claims.

### Offline workflow reconstruction

Use when the original task reads databases/APIs or writes to mail, calendars, Google Sheets, or similar services.

- Extract only fields and rows needed by the original task from the same immutable fixture.
- Retain the extraction script, fixture revision/checksum, query or projection logic, and excluded sensitive fields.
- Freeze relative time as an explicit snapshot date supported by the upstream state.
- Replace an external side effect with a local equivalent that preserves the same information and decision, such as an email draft, schedule sheet, or standalone workbook.
- Do not add or repair records to make the task more convenient.

### Evidence-backed composition

Composition is allowed only when all sources are compatible parts of one plausible request. Assign one role to each source:

- **task source:** supplies the user intent and workflow;
- **input source:** supplies task-relevant files or state;
- **rule source:** supplies domain rules or acceptance constraints;
- **presentation reference:** supplies an explicitly visible style target;
- **evaluation reference:** supplies a checking method, not hidden truth.

The task source must remain identifiable. Inputs from another source may replace unavailable assets only when their schema and business meaning fit without inventing facts, their license permits use, and the adaptation record explains the substitution. Do not join unrelated finance, HR, and charting tasks to force an L4 label. Do not transfer facts or labels from one organization's dataset into another source's business context without explicit normalization and human review.

## Conflict resolution

When sources disagree, apply this order:

1. explicit user request;
2. facts and state in user-visible inputs;
3. explicit domain rules supplied with the task;
4. deterministic upstream evaluator semantics;
5. upstream Golden;
6. author preference.

Document every material conflict. An upstream Golden may be corrected when it changes unrequested fields, invents figures, mistranslates values, or contradicts the input. Preserve the flawed upstream artifact as audit evidence when licensing permits, and make the local Judge reject the unsupported behavior where practical.

## Leakage and source diversity

Assign a leakage group before rollout using source task, input hashes, template, semantic workflow, and derived state. Variants sharing a workbook, source fixture, formula family, screenshot, or task template belong together even if wording differs. Candidate counts do not equal usable diversity when many rows share one evidence base.
