# Build the CS161 roster verification register

Use the frozen `CS161_Class_Roster.xlsx` workbook to create exactly one editable Word document at `/app/output/CS161/students.docx`. This register will let course operations staff trace every listed CS161 student back to the source row, review mechanical score exceptions, and prove that the long roster was transferred without omission or reordering.

## Canonical roster

- Include all 123 source rows in worksheet order. For every student, preserve Source row, Student ID, Last Name, First Name, Academic Standing, HW1, HW2, Midterm, HW3, and Final exactly.
- Divide the source register into five labeled blocks with these inclusive worksheet ranges: B1 `2–26`, B2 `27–51`, B3 `52–76`, B4 `77–101`, and B5 `102–124`.
- Show each block's row count, first and last Student ID, and a reproducible short digest in a closing block ledger. Do not alphabetize, drop, merge, or duplicate source rows.

## Source-driven reconciliation

- Provide an overall profile for each of the five assessment columns: arithmetic mean and median to two decimals, minimum, maximum, count of exact zeros, and count of exact 100s.
- Reconcile Freshman, Sophomore, Junior, and Senior in that order. Show each standing's student count, five assessment means to two decimals, and count of students included in the verification queue.
- Provide a standing-by-rule incidence table and an all-student total.

## Mechanical verification queue

Evaluate every student using these four rules, based only on the five source scores:

- `ZERO`: at least one score equals 0.
- `PERFECT`: at least one score equals 100.
- `SPREAD80`: maximum score minus minimum score is at least 80.
- `SWING60`: absolute value of Final minus Midterm is at least 60.

Include only students who hit at least one rule. Sort by number of rule hits descending, then assessment spread descending, then source row ascending. For every queued student show rank, source row, Student ID, `First Name Last Name`, hit count, unweighted five-score mean to one decimal, spread, signed `Final − Midterm`, and the applicable rule codes in the order above. Split the queue into ranks `1–20`, `21–40`, and `41–56`, and reconcile those partitions on the closing page.

## Word structure and control boundary

- Organize the document as 11 controlled pages: a portrait control/navigation page; one landscape profile page; five landscape roster-block pages; three landscape queue pages; and a portrait closure page.
- Use native editable Word tables, heading styles, repeating table headers, rows that do not split across pages, role-appropriate headers, and `PAGE / NUMPAGES` footer fields.
- From the control page, provide working internal links to the profile, five roster blocks, three queue parts, and closure. Use matching Word bookmarks at those destinations.
- Record these custom properties: `SourceSHA256`, `RosterCount`, `VerificationCount`, `FirstStudentID`, `LastStudentID`, `SourceSheet`, and `RulesetVersion`.
- Keep the document offline-safe: no macros, embedded files, external links, or page images used as a substitute for editable content. Deliver no other DOCX.
- State clearly that the flags are mechanical source checks only. Do not infer a course grade, academic consequence, misconduct finding, or advising decision.
