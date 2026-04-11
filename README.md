# Validation Report 0001 — Excel Auditor

A static web app for auditing IE vs Field Nation Excel reports side by side.

## Deploy to GitHub Pages

1. Upload all files to a GitHub repository
2. Go to **Settings → Pages**
3. Set source to **Deploy from a branch** → `main` → `/ (root)`
4. Save — your site will be live at `https://<username>.github.io/<repo-name>/`

## Features
- Upload Excel workbook → auto-parses the Report sheet
- Side-by-side IE vs FN comparison with mismatch highlighting
- ITD exception rule, date/time/name validation
- Live charts and stats dashboard
- Manual "Correct" tracker with localStorage persistence
- Export to XLSX (5 sheets) or Outlook .eml
