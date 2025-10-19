# Our Home Reporting Portal (static site)

This folder contains a production‑ready static site plus a small Python helper to pull **Last Updated** dates from Azure SQL for the Power BI dashboards.

## Files
- `index.html` – main page (Bootstrap 5, search, filters, sort, responsive grid)
- `assets/css/style.css` – theme styles
- `assets/js/data.js` – report catalog (name, type, link, logo). **Edit this** to add/remove dashboards.
- `assets/js/app.js` – renders cards and loads `last_updated.json`
- `last_updated.json` – generated map: `slug -> YYYY-MM-DD` (used only for Power BI entries)
- `fetch_last_updated.py` – connects to Azure SQL and fills `last_updated.json` (uses fuzzy matching to map report names to tables like `dbo.[Our Home *]`)

## Usage
1. Upload the whole folder to your web root (e.g., `/var/www/html/our-home-portal/`).
2. Place retailer logos in `assets/logos/` (optional; placeholders included). Keep the filenames referenced in `data.js`.
3. (Optional) Run the updater from the server (with Python 3, `pyodbc`, `tenacity`, and ODBC Driver 17 installed):

```bash
cd our-home-portal
python3 -m pip install pyodbc tenacity
python3 fetch_last_updated.py
```

This creates/updates `last_updated.json`. The site will automatically display dates when present.

### Scheduling (optional)
Use `cron` to refresh once a day:
```
0 6 * * *  cd /var/www/html/our-home-portal && /usr/bin/python3 fetch_last_updated.py
```

## Security
Links are public Power BI view URLs you provided. If you later embed private/secured reports, no code changes are needed—just update the URL in `data.js`.

---
