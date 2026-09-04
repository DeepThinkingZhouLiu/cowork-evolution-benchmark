# officebench-2-39-all-emails

## Original OfficeBench request

> find all Bob's emails in this year, save them in emails.docx by ascending order

## Delivery requirements

Complete the request using the visible inputs and deliver only the editable DOCX file(s) `/app/output/emails.docx`. The fixed output path is an offline Harbor adaptation. Preserve the source facts, ordering, and requested Word content; do not create non-DOCX deliverables.

## Task-specific details

Compile all four supplied messages into editable `/app/output/emails.docx`. The files have no `Date` header, so use the visible `Sent at` timestamp in each message body as the chronology field and sort strictly ascending by that value. Include exactly one native Word table with the columns `Sent at`, `From`, `Subject`, and `Body`; preserve every field and full body without adding external information, and briefly state the chosen date basis.
