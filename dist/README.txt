Validation Report Auditor v8

This version is prepared for GitHub Pages deployment.

Main fixes:
- index.html is at project root
- asset paths are relative
- GitHub Actions workflow included
- date parsing avoids timezone / day offset issues by reading Excel display values as date-only strings
- runtime banner added for load errors instead of silent white screen behavior

Commands:
- npm run check
- npm run build
- python3 -m http.server 4173

Deploy output:
- dist/
