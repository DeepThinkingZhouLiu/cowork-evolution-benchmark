# officebench-3-84-sleep-analysis

## Original OfficeBench request (not executable against the released input)

> Analyze relationship between sleeping hours and age, weight, gender from survey result, convert report to report.docx file

## Delivery requirements

Complete the request using the visible inputs and deliver only the editable DOCX file(s) `/app/output/report.docx`. The fixed output path is an offline Harbor adaptation. Preserve the source facts, ordering, and requested Word content; do not create non-DOCX deliverables.

## Task-specific details

The released workbook contains `Age`, `Gender`, `Income`, and `Satisfaction Score`, but no sleep-hours or weight columns. For this offline adaptation, analyze how `Satisfaction Score` relates to age, income, and gender in `survey_responses.xlsx`. Report the row count, overall mean, Pearson correlations with age and income, and mean satisfaction by gender in editable `/app/output/report.docx`. Do not invent unavailable fields.
