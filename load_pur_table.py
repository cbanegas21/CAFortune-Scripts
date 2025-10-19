"""
load_pur_table.py
• Pulls Repsly pages starting at START_ID.
• Keeps only the forms whose cleaned name is in FORM_WHITELIST.
• Streams rows into dbo.<form_name>, adding columns on-the-fly.
• Auto-shrinks over-long column names so SQL identifiers ≤128 chars.
• Uses driver 17, 5-second SQL timeouts, unlimited input buffers.
"""

import logging, re, requests, pandas as pd, pyodbc, hashlib
from datetime import datetime
from contextlib import closing
from requests.auth import HTTPBasicAuth
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

# ─────── CONFIG ────────────────────────────────────────────
SERVER   = "ca-data-server.database.windows.net"
DATABASE = "CAFortuneDatabase"
USERNAME = "sqladmin"
PASSWORD = "Maxine2021."
DRIVER   = "{ODBC Driver 17 for SQL Server}"       # driver 17

API_USER = "80941603-F785-4E0F-8AB1-ED798E54F88C"
API_PSW  = "17BCD0FD-94DD-4B0C-B059-76D68C1145A8"
API_URL  = "https://api.repsly.com/v3/export/forms/"

START_ID = 243_588_639         # first page to fetch
FORM_WHITELIST = [
    "pur_gum_the_fresh_market",
    "pur_gum_sprouts",
    "pur_gum_whole_foods"
]

CHUNK_SIZE   = 1000
CONN_TIMEOUT = 5                 # shorter → quick retry
MAX_DB_TRY   = 5
API_TIMEOUT  = 30

# ─────── LOGGING ───────────────────────────────────────────
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s | %(levelname)-8s | %(message)s")
log = logging.getLogger(__name__)

# ─────── HELPERS ───────────────────────────────────────────
clean = lambda t: re.sub(r"[^0-9a-z_]", "_", t.lower().replace(" ", "_"))
ms_dt = lambda s: datetime.utcfromtimestamp(int(s[6:19]) / 1000) if s.startswith("/Date(") else None

MAX_ID_LEN = 128                 # SQL Server hard limit
HEAD_LEN   = 120                 # leave room for “_xxxxxxx” hash

def shrink(name: str) -> str:
    """Trim identifiers >128 chars and append a 7-char hash for uniqueness."""
    if len(name) <= MAX_ID_LEN:
        return name
    head = name[:HEAD_LEN]
    tail = hashlib.md5(name.encode()).hexdigest()[:7]
    return f"{head}_{tail}"

def wide_buffers(cur, ncols):
    """Allocate unlimited buffer for every column; safe to call each insert."""
    cur.setinputsizes([(pyodbc.SQL_WVARCHAR, 0, 0)] * ncols)
    cur.fast_executemany = True

@retry(
    stop=stop_after_attempt(MAX_DB_TRY),
    wait=wait_fixed(CONN_TIMEOUT),
    retry=retry_if_exception_type(pyodbc.Error),
    before_sleep=lambda r: log.warning("DB login failed – retrying (%s/%s)…",
                                       r.attempt_number, MAX_DB_TRY),
)
def open_conn():
    log.info("  ↳ SQL login …")
    return pyodbc.connect(
        f"DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};"
        f"UID={USERNAME};PWD={PASSWORD};Encrypt=yes;TrustServerCertificate=no;"
        f"Connection Timeout={CONN_TIMEOUT};LoginTimeout={CONN_TIMEOUT};",
        autocommit=False,
    )

def ensure_table(cur, tbl, cols):
    cur.execute("IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name=?) "
                "EXEC ('CREATE TABLE ['+?+'] (" +
                ",".join(f'[{c}] NVARCHAR(MAX)' for c in cols) + ")')",
                tbl, tbl)

def add_missing(cur, tbl, cols):
    cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME=?", tbl)
    have = {r.COLUMN_NAME.lower() for r in cur}
    for c in cols:
        if c.lower() not in have:
            cur.execute(f"ALTER TABLE [{tbl}] ADD [{c}] NVARCHAR(MAX)")
            cur.commit()
            log.info("  • columna %s agregada a %s", c, tbl)

def insert_chunk(cur, tbl, df, cols):
    sql = f"INSERT INTO [{tbl}] ({', '.join(f'[{c}]' for c in cols)}) VALUES ({', '.join('?'*len(cols))})"
    wide_buffers(cur, len(cols))
    safe = df[cols].astype(object).where(pd.notnull(df), None)
    for start in range(0, len(safe), CHUNK_SIZE):
        cur.executemany(sql, safe.iloc[start:start+CHUNK_SIZE].values.tolist())
        cur.commit()

# ─────── MAIN ──────────────────────────────────────────────
log.info("=== PUR loader starting (from ID %s) ===", START_ID)
rows_kept, last_id = [], START_ID

while True:
    r = requests.get(f"{API_URL}{last_id}",
                     auth=HTTPBasicAuth(API_USER, API_PSW),
                     timeout=API_TIMEOUT)
    r.raise_for_status()
    page = r.json().get("Forms", [])
    if not page:
        break

    for f in page:
        if FORM_WHITELIST and clean(f["FormName"]) not in FORM_WHITELIST:
            continue
        row = {
                "formid":            f["FormID"],
                "clientcode":        f["ClientCode"],
                "clientname":        f["ClientName"],
                "dateandtime":       ms_dt(f["DateAndTime"]),
                "representativecode": f["RepresentativeCode"],
                "representativename": f["RepresentativeName"],
                "streetaddress":     f["StreetAddress"],
                "zip":               f.get("Zip") or f.get("ZIP"),
                "city":              f["City"],
                "state":             f["State"],
                "country":           f["Country"],
                "email":             f["Email"],
                "phone":             f["Phone"],
                "mobile":            f["Mobile"],
                "territory":         f["Territory"],
                "longitude":         f["Longitude"],
                "latitude":          f["Latitude"],
                "signatureurl":      f["SignatureURL"],
                "visitstart":        f["VisitStart"],      # ← pon ms_dt() si lo necesitas en datetime
                "visitend":          f["VisitEnd"],        # ← idem
                "visitid":           f["VisitID"],
            }

        for it in f.get("Items", []):
            col_name = shrink(clean(it["Field"]))
            row.setdefault(col_name, it.get("Value"))
        rows_kept.append((clean(f["FormName"]), row))

    nxt = r.json()["MetaCollectionResult"]["LastID"]
    log.info("  • page done – next ID %s   (kept rows so far %s)", nxt, len(rows_kept))
    if nxt <= last_id:
        break
    last_id = nxt

if not rows_kept:
    log.info("No matching forms found – exiting.")
    raise SystemExit

with closing(open_conn()) as conn:
    cur = conn.cursor()
    tables = {}
    for tbl, row in rows_kept:
        tables.setdefault(tbl, []).append(row)

    for tbl, rows in tables.items():
        df = pd.DataFrame(rows)
        # shrink any over-long column names before touching SQL
        df.columns = [shrink(c) for c in df.columns]
        cols = list(df.columns)

        ensure_table(cur, tbl, cols)
        add_missing(cur, tbl, cols)
        insert_chunk(cur, tbl, df, cols)
        log.info("✓ inserted %s rows into %s", len(df), tbl)

log.info("=== FINISHED: total rows inserted %s ===", len(rows_kept))
