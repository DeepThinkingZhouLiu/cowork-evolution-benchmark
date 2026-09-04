# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: 8559f5246583014288b3c81d98892874c7610c3586e95799035c6f77d571fd63
- E4 `Alice.ics` contains 7 VEVENT blocks. `SUMMARY` counts recompute to `Workout=3`, `Class=2`, and `Meeting=2`; `Workout` is the unique maximum.
- E5 Forbidden source terms: none
- E6 Deliverable contract: events.docx
- E7 Golden: `solution/` is audit-only; the Judge parses all source `SUMMARY` values and does not use Golden bytes as truth.
