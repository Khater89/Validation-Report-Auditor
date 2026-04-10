Validation Report Auditor v5

What is new in this version:
- Specific Date filter at the top of the interface
- From / To date range filter for week or month analysis
- Strong Status Filter with Errors only / Corrected only / Matched only / Matched by exception
- Cards now reflect the current visible filtered set
- Exported workbook includes both full data and the current filtered view
- Outlook draft summary now reflects the current filtered view

How to use:
1) Open index.html in your browser.
2) Upload the Excel workbook.
3) Use the top filters:
   - Specific Date: one exact day
   - From Date / To Date: date range analysis
   - Status Filter: for example Errors only
4) Review rows.
5) Click Correct for any error already fixed outside the app.
6) Export Final XLSX or Outlook Email.

Date filter logic:
- The app uses FN Scheduled Date first.
- If FN Scheduled Date is empty, it uses IE Due Date.

Validation logic:
- IE & FN Status must be OK.
- IE Due Date must match FN Scheduled Date.
- IE Due Time must match FN Scheduled Time.
- IE FE Name must match FN FE Name.
- ITD: date mismatch is accepted.
- ITD or ITR: time mismatch is accepted.

Tracking logic:
- Correct does not edit data inside the app.
- It only marks the row as handled externally and changes the status to Corrected.


Date offset fix:
- Date cells that arrive from Excel as browser Date objects are now rendered from local date parts only.
- Numeric Excel serial dates are converted without timezone shifting.
- This prevents the one-day-back issue and preserves the same date shown in Excel.
