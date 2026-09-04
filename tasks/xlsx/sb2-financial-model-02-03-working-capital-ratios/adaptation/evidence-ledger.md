# Evidence ledger

| Evidence | Immutable source | Local path | Role | License | Transformation | SHA256 | Supports | Risk |
|---|---|---|---|---|---|---|---|---|
| E1 | SpreadsheetBench-2 sample `Financial_Model/02_03` at `d8a14da1898e48d07c8a9ff8974d0effc9302ece` | `data/input_files/02_03_input.xlsx` | Solver-visible input | CC-BY-SA-4.0 | Original bytes | `b074b65f57ff82918a5ac56b79a39ebfc3ca937ddf674b6834642b1fc2fc2634` | Workbook state, labels, neighboring formula patterns | Upstream fixture is synthetic |
| E2 | Same dataset record | `adaptation/upstream-task.json` | Original request | CC-BY-SA-4.0 | Extracted record | n/a | Requested metrics and repair rules | Upstream wording audited below |
| E3 | Same dataset Golden | `adaptation/upstream-golden.xlsx` | Audit-only reference | CC-BY-SA-4.0 | Original bytes | `d2b71c081d3fb124ccf41ebe357f7b92439bab93b3e0cb77279e47289ca3862f` | Cross-check of independently justified formulas and structure | Not read by Judge |
| E4 | Dataset card and retained license | `adaptation/LICENSE-CC-BY-SA-4.0.txt` | Rights evidence | CC-BY-SA-4.0 | Retained text | `28a9529c7d0bb4dc51f4bf5c116a3d16ef247a052f7591466768ddf563fd1cf5` | Redistribution and ShareAlike basis | Human legal review remains required |

Every instruction requirement maps to E1+E2. Formula checks use visible labels, neighboring formulas and stated business rules; preservation and integrity checks use E1. E3 is used only to audit disagreements and the reference solution.
