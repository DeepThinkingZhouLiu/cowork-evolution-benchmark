# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: 8e39f62f9fb5747c5ac1527eea0162ef7f91a5f178a9890eb46a995ef5d5207c
- E4 Required source terms: Market, Analysis, Marketing, Plan, Business, Customer, How Large
- E5 Forbidden source terms: none
- E6 Deliverable contract: market_analysis.docx
- E7 Golden: `solution/` is audit-only; the Judge recomputes from E1-E3 and does not use Golden bytes as truth.
