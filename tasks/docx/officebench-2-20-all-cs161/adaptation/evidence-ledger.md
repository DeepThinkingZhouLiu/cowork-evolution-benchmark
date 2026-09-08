# CS161 roster control book evidence ledger

- E1 Query / evaluator parameters: `adaptation/source-task.json`, SHA256 `402315cb9ee0f0907e978c97ecd432d6b0c5bce45f23fb10d85f6b47d5ee9e62`.
- E2 Upstream evaluator implementation: OfficeBench `utils/evaluate.py`, commit `b978b808667c32b52ce19a67ce1def1de9ae02b7`, SHA256 `f4a675a590a57844fd21d6d6660a00d8c69661b0f92dc26146f9b8869a26cc90`; audit excerpt `adaptation/upstream-evaluator.py`.
- E3 License: `adaptation/LICENSE-OfficeBench`, Apache-2.0, SHA256 `c71d239df91726fc519c6eb72d318ec65820627232b2f796219e87dcf35d0ab4`.
- E4 Workbook bytes: `CS161_Class_Roster.xlsx`, SHA256 `c3289e2f535d113836e14f0a10460361431749582f6e336931fbe9fcb6bc9ed7`; upstream/data/environment copies match byte-for-byte.
- E5 Entity ledger: 9 populated records at `Sheet1!2:10`; all ten source columns are transcribed into the roster register.
- E6 Aggregation ledger: Major counts `CS=3, ECE=2, Math=3, Stats=1`; standing counts `Freshman=4, Junior=2, Senior=2, Sophomore=1`.
- E7 Assessment ledger: `HW1 9/30.44/6/58`; `HW2 9/58.11/3/94`; `Midterm 9/53.44/6/88`; `HW3 9/50.44/0/96`; `Final 9/53.89/1/100` (N/average/min/max).
- E8 Data-quality ledger: 9/9 complete rows; 9/9 unique IDs; 9/9 IDs match `S` plus four digits; 45/45 integer scores lie in `[0,100]`.
- E9 Navigation ledger: one `roster_<Student ID>` bookmark per source-order row and one index hyperlink per alphabetically sorted student.
- E10 Deliverable migration: `students.docx` from v0.2 is superseded by `cs161_roster_control_book.docx` in v0.3; no non-DOCX output is required.
- E11 Golden: `solution/cs161_roster_control_book.docx`, SHA256 `325bb9b8a1dfcb5f3763d0e4f9583a8a8671fe4d84202a3d5a68460274dd912c`; Judge recomputes E5-E9 from E4 and never compares Golden bytes.
