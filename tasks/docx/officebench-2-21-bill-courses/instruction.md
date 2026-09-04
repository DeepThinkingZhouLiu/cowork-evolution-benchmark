# officebench-2-21-bill-courses

## Original OfficeBench request

> Extract Bill's previous courses from Bill_course_list.xlsx, and put the course in classes_taken.docx. Each course should be on a seperate line.

## Delivery requirements

Complete the request using the visible inputs and deliver only the editable DOCX file(s) `/app/output/classes_taken.docx`. The fixed output path is an offline Harbor adaptation. Preserve the source facts, ordering, and requested Word content; do not create non-DOCX deliverables.

## Task-specific details

Read `Bill_course_lists.xlsx` and place every value in its `Course` column into `/app/output/classes_taken.docx`, one editable paragraph per course in source order. Keep course names exactly as shown and do not include other file formats.
