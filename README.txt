Validation Report Project - Updated Exception Logic

Files included:
- index.html
- validation_report_04022026_custom_v2.xlsx
- README.txt

How to use:
1) Open index.html in a browser.
2) Upload the source workbook that contains the Report sheet.
3) The app will validate only rows where Assigned Company = Field Nation.

Applied rules:
- IE & FN Status must be OK.
- IE Due Date must match FN Scheduled Date.
- IE Due Time must match FN Scheduled Time.
- IE FE Name must match FN FE Name after normalization.
- Date mismatch is accepted only when Ticket Type = ITD.
- Time mismatch is accepted when Ticket Type = ITD or ITR.

Updated result on the provided file:
- Comparable Tickets: 233
- Matched: 185
- Errors: 48
- Completion: 79.4%
- Time Exceptions Applied: 79
  - ITD Time Exceptions: 78
  - ITR Time Exceptions: 1
