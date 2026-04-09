Validation Report Auditor v4
============================

What changed in this version
----------------------------
- No pop-up correction window.
- Every open error row shows a "Correct" button.
- After you fix the source externally, click "Correct" in the app.
- The button changes immediately to "Corrected" and becomes green.
- The row is counted as Corrected in statistics and export.
- Export includes summary statistics, final report, corrected rows, and remaining open errors.
- Outlook draft export (.eml) is included.

How to use
----------
1. Open index.html in your browser.
2. Upload the Excel workbook.
3. Review rows with Final Result = Error.
4. Fix the source outside the app if needed.
5. Return to the app and click Correct for each handled row.
6. Export the final XLSX report or Outlook email draft.

Files
-----
- index.html
- app.js
- styles.css
- sample_input.xlsx

Notes
-----
- The app works best with a workbook that contains a sheet named Report.
- Tracking state is saved in browser localStorage per uploaded file.
