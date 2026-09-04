# Evidence ledger

- E1 Query: `adaptation/source-task.json` field `task`
- E2 Evaluation: `adaptation/source-task.json` field `evaluation`
- E3 Input bytes: `metadata/input-manifest.json`; recomputed package hashes: d87d482d5ae7960e2e43d7dd6d21377e60e73e8fce1bf2a01aff7aca8a08c537
- E4 PDF pp. 1-5: title/authors, the all-attention contribution, 6-layer encoder and decoder, `d_model=512`, `d_ff=2048`, 8 heads, `d_k=d_v=64`, decoder masking, and sinusoidal positional encoding.
- E5 PDF pp. 6-9: self-attention complexity and path comparison; WMT corpus sizes; 8 P100 GPUs; 100,000-step/12-hour base training and 300,000-step/3.5-day big training; 28.4 and 41.0 BLEU results.
- E6 PDF p. 9: conclusion and future work on other modalities, restricted attention, and less sequential generation.
- E7 Deliverable contract: `paper.docx`, with native headings, paragraphs, and tables.
- E8 Golden: `solution/` is audit-only; the Judge extracts anchors from the fixed PDF and does not use Golden bytes as truth.
