# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: e6468e5452830bc4bd8767fb9a95c0ea95a65d876e306581437889e43c2a821a
- E4 Required source terms: apple, cabbage, onion
- E5 Forbidden source terms: none
- E6 Deliverable contract: shopping_list.docx
- E7 Golden: `solution/` is audit-only; the Judge recomputes from E1-E3 and does not use Golden bytes as truth.
