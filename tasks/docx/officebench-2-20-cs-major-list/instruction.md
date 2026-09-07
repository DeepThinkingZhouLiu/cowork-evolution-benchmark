# CS Major Advising Review Docket

The academic review coordinator needs an evidence-controlled Word docket for the students whose `Major` is exactly `CS` in `CS161_Class_Roster.xlsx`. Create and deliver only the editable file `/app/output/cs_major_advising_docket.docx`.

Use the workbook as the sole factual source. Treat the first worksheet as the source register, keep names as `First Name Last Name`, and use one decimal place for every calculated average or average difference. Do not expose any non-CS student's name, Student ID, standing, or individual score in the docket.

Build the document in this order:

1. A title area and source-control statement naming `CS161_Class_Roster.xlsx`, followed by a **Review Queue** in a native Word table. Include Review order, Student ID, Name, Standing, Final, Midterm -> Final, Review tier, Source row, and an internal **Open evidence record** link. Rank the CS students by tier, then Final ascending, then Student ID ascending.
2. Apply this exact review-tier rule: Final `0-5` = **Priority 1**; Final `6-20` = **Priority 2**; Final `21-100` = **Priority 3**. A Midterm -> Final value is `Final - Midterm` and must carry its sign. Attach exactly one native Word comment to each Review tier value in the queue. Use the comment text `Tier rule: Final 0-5 -> Priority 1. Observed Final: N.`, `Tier rule: Final 6-20 -> Priority 2. Observed Final: N.`, or `Tier rule: Final 21-100 -> Priority 3. Observed Final: N.`, as applicable.
3. Add one separately paged **Evidence Record** for each selected student, in review-queue order. Each record heading must be a bookmark targeted by its queue link. Include an editable identity table; an assessment table containing HW1, HW2, Midterm, HW3, Final, each exact score and its worksheet cell locator; and a derived-checks table containing five-assessment average, lowest assessment and score, highest assessment and score, score range, Midterm -> Final, trend (`Decline`, `Flat`, or `Increase`), and review tier.
4. Finish on a new page with a **Cohort Comparison Appendix**. For HW1, HW2, Midterm, HW3, and Final, show CS N and average, non-CS N and average, and `CS average - non-CS average`. This appendix may contain non-CS aggregates only, never non-CS identities.
5. Add a **Control Summary** showing source row count, CS selected count, non-CS excluded count, the three selected source rows, the tier rules, the ranking rule, and a reconciliation statement confirming that the queue and evidence-record counts match.

Use native editable Word text and tables, Heading styles, repeating table headers, non-splitting data rows, a consistent page header, and a footer with an automatic page-number field. Keep every page portrait, keep the five required pages readable without blank pages or clipped tables, and do not use screenshots, flattened page images, macros, embedded files, OLE objects, external links, or extra deliverables.
