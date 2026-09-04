# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: d7764b78e2931fbcda87f2a45ef140693c3d6977e0235b264aa868cf949bab8a
- E4 Required source terms: CS 101, MATH 150, PHYS 101, CS 102, MATH 151, PHYS 102, CS 103, MATH 152, PHYS 103, CS 104, MATH 250, PHYS 104, CS 105, MATH 251, PHYS 105
- E5 Forbidden source terms: none
- E6 Deliverable contract: classes_taken.docx
- E7 Golden: `solution/` is audit-only; the Judge recomputes from E1-E3 and does not use Golden bytes as truth.
