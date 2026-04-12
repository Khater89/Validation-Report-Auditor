# Validation Report 0001 — Excel Auditor
## GitHub Pages Deployment Guide

### Files
| File | Purpose |
|------|---------|
| `index.html` | The entire application (self-contained) |
| `404.html` | Copy of index — prevents white screen on direct URL access |
| `README.md` | This file |

### How to Deploy

**Step 1 — Create a GitHub Repository**
1. Go to https://github.com/new
2. Name it e.g. `validation-report-auditor`
3. Set to Public (required for free GitHub Pages)
4. Click **Create repository**

**Step 2 — Upload Files**
Option A (Web UI):
1. Click **Add file → Upload files**
2. Drag all 3 files into the upload area
3. Commit: "Initial deployment"

Option B (Git CLI):
```bash
git init
git add .
git commit -m "Initial deployment"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/validation-report-auditor.git
git push -u origin main
```

**Step 3 — Enable GitHub Pages**
1. Go to your repo → **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` → `/ (root)` → **Save**
4. Wait ~60 seconds, then visit:
   `https://YOUR_USERNAME.github.io/validation-report-auditor/`

### Why This Works on GitHub Pages
- ✅ No build step / no Node.js needed
- ✅ All libraries loaded from CDN (SheetJS + Chart.js) with fallback mirrors
- ✅ No relative file paths that could break
- ✅ 404.html prevents white screen on page refresh
- ✅ Excel dates decoded with UTC math (no timezone Day-1 offset)
- ✅ Debug overlay shows detailed errors if something goes wrong in production

### Troubleshooting
If the app loads but Excel file doesn't process:
1. Open browser DevTools (F12) → Console tab
2. Check for red errors
3. The debug overlay will appear automatically with details
4. Most common cause: ad-blocker blocking CDN requests
   → Disable extension for your GitHub Pages domain

### Excel File Requirements
- Format: `.xlsx` (Excel 2007+)
- Must contain a sheet named **Report**
- Column A: Ticket number
- Column F (index 5): Ticket Type (e.g. ITD, ITR)
- Column M (index 12): Assigned Company — must equal `Field Nation`
- Columns H,I (7,8): IE Due Date, IE Due Time
- Columns U,V (20,21): FN Scheduled Date, FN Scheduled Time
- Column N,X (13,23): IE FE Name, FN FE Name
- Column AA (26): IE & FN Status
