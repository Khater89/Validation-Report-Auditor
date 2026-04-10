# Validation Report Auditor v8

GitHub Pages-ready static web app for validating Excel report rows between IE and Field Nation.

## What was fixed in this version
- Root-ready file structure for GitHub Pages (`index.html` at the project root)
- Relative asset paths (`./styles.css`, `./app.js`)
- Clean GitHub Actions workflow for Pages deployment
- Runtime banner instead of silent failure when the Excel library or workbook load fails
- Safer Excel date handling: reads formatted Excel values as **date-only strings** to avoid timezone day-offset issues
- Retains all required features:
  - Specific Date filter
  - Date Range filter (From / To)
  - Errors Only filter
  - Correct / Corrected tracking button
  - ITD / ITR exception rules
  - Final XLSX export
  - Outlook `.eml` export

## Local run
Open `index.html` directly, or use a tiny local server:

```bash
python3 -m http.server 4173
```

Then open `http://localhost:4173`

## Build
```bash
npm run check
npm run build
```

The deployable output is created in `dist/`.

## GitHub Pages deployment
This repo includes a GitHub Actions workflow at:

`.github/workflows/deploy.yml`

It will:
1. validate `app.js`
2. build `dist/`
3. deploy `dist/` to GitHub Pages

## Important deployment note
This is a **static app**, not React. It does **not** need React Router, Vite, or BrowserRouter/HashRouter fixes.
