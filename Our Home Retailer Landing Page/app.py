# app.py — serve the site + an in-site API to pull MAX(DateAndTime)
import os, re, json
from flask import Flask, request, send_from_directory, jsonify
import pytds  # pure-Python SQL Server driver (no ODBC)

# Read creds from environment (Azure App Service > Configuration)
SQL_SERVER   = os.getenv("SQL_SERVER", "ca-data-server.database.windows.net")
SQL_DATABASE = os.getenv("SQL_DATABASE", "CAFortuneDatabase")
SQL_USERNAME = os.getenv("SQL_USERNAME", "sqladmin")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "YOUR_PASSWORD")  # set in env on Azure

app = Flask(__name__, static_folder=".", static_url_path="")

# Serve index and static files
@app.route("/")
def index():
    return send_from_directory(".", "index.html")

# API: POST /api/refresh-dates  { reports: [{name, slug}, ...] }
@app.route("/api/refresh-dates", methods=["POST"])
def refresh_dates():
    body = request.get_json(silent=True) or {}
    reports = body.get("reports") or []
    out = {}

    # Connect securely (TLS)
    with pytds.connect(SQL_SERVER, database=SQL_DATABASE, user=SQL_USERNAME, password=SQL_PASSWORD,
                       port=1433, tds_version=7.4, use_tz=False, timeout=30,
                       auth='sql', encrypt=True, validate_host=False) as conn:
        with conn.cursor() as cur:
            for r in reports:
                name = r.get("name")
                slug = r.get("slug")
                if not name or not slug:
                    continue
                tbl = find_best_table(cur, name)
                if not tbl:
                    continue
                if not has_dateandtime(cur, tbl):
                    continue
                dt = max_date(cur, tbl)
                if dt:
                    out[slug] = dt.strftime("%Y-%m-%d")
    return jsonify(out)

# --- helpers ---
def norm(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', (s or "").lower())

def find_best_table(cur, report_name: str):
    m = re.search(r'(?:OUR HOME|Our Home)\s+(.*)', report_name or "")
    base = m.group(1) if m else (report_name or "")
    key = norm(base)
    cur.execute("SELECT name FROM sys.tables WHERE name LIKE 'Our Home %'")
    rows = [row[0] for row in cur.fetchall()]
    best = None; best_score = -1
    for t in rows:
        tn = norm(t.replace('Our Home','').strip())
        score = 0
        if key in tn or tn in key: score += 2
        toks1=set(re.findall(r'[a-z0-9]+', key)); toks2=set(re.findall(r'[a-z0-9]+', tn))
        score += len(toks1 & toks2)
        if score > best_score: best_score, best = score, t
    return best if best_score > 0 else None

def has_dateandtime(cur, table: str) -> bool:
    cur.execute("""
        SELECT 1
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=%s AND COLUMN_NAME='DateAndTime'
    """, (table,))
    return cur.fetchone() is not None

def max_date(cur, table: str):
    cur.execute(f"SELECT CONVERT(date, MAX([DateAndTime])) FROM [dbo].[{table}]")
    row = cur.fetchone()
    return row[0] if row else None

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
