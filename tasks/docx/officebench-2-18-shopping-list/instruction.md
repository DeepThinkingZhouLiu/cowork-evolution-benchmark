# officebench-2-18-shopping-list

## Original OfficeBench request

> write all shopping item in docx file (shopping_list.docx) from shopping list (shopping_list.xlsx)

## Delivery requirements

Complete the request using the visible inputs and deliver only the editable DOCX file(s) `/app/output/shopping_list.docx`. The fixed output path is an offline Harbor adaptation. Preserve the source facts, ordering, and requested Word content; do not create non-DOCX deliverables.

## Task-specific details

Read every item and amount from `shopping_list.xlsx` and create `/app/output/shopping_list.docx`. Use an editable Word table with the same item names, amounts, and source order. Do not omit or invent rows.
