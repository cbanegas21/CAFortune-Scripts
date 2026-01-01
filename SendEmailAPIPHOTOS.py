import re, time, datetime as dt, sys
from pathlib import Path
import requests, pandas as pd, openpyxl
from requests.auth import HTTPBasicAuth

# ── CONFIG ──────────────────────────────────────────────────────────
START_DATE = dt.date(2025, 12, 12)       # inclusive
END_DATE   = dt.date(2025, 12, 27)       # inclusive

REPSLY_USER = "80941603-F785-4E0F-8AB1-ED798E54F88C"
REPSLY_PASS = "17BCD0FD-94DD-4B0C-B059-76D68C1145A8"

BASE_URL = "https://api.repsly.com/v3/export/photos/"
HEADERS  = {"Accept": "application/json"}

# ── HELPERS ─────────────────────────────────────────────────────────
if START_DATE > END_DATE:
    sys.exit("ERROR: START_DATE must be on or before END_DATE.")

T0 = dt.datetime.combine(START_DATE, dt.time.min)
T1 = dt.datetime.combine(END_DATE,   dt.time.max)

def parse_repsly_date(raw):
    m = re.search(r"/Date\((\d+)", raw or "")
    return dt.datetime.utcfromtimestamp(int(m.group(1)) / 1000) if m else None

def get_photo_url(photo_dict):
    """Return whichever key actually exists."""
    return (
        photo_dict.get("PhotoURL")   # most common
        or photo_dict.get("PhotoUrl")
        or photo_dict.get("Url")
        or ""
    )

# ── FETCH & FILTER ──────────────────────────────────────────────────
print(f"Fetching tagged photos {START_DATE} -> {END_DATE} …")
photos = []
last_id = 0

while True:
    resp = requests.get(
        f"{BASE_URL}{last_id}",
        headers=HEADERS,
        auth=HTTPBasicAuth(REPSLY_USER, REPSLY_PASS),
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    batch = data.get("Photos", [])
    if not batch:
        break

    for p in batch:
        tag = (p.get("Tag") or "").strip()
        dt_parsed = parse_repsly_date(p.get("DateAndTime"))
        if tag and dt_parsed and T0 <= dt_parsed <= T1:
            photos.append(
                {
                    "tag": tag,
                    "photo_url": get_photo_url(p),
                    "date": dt_parsed.date().isoformat(),
                }
            )

    last_id = data.get("MetaCollectionResult", {}).get("LastID") or last_id
    if not last_id:
        break
    time.sleep(0.3)

print("Rows kept:", len(photos))
if not photos:
    sys.exit("No tagged photos in that date range.")

# ── SAVE EXCEL ──────────────────────────────────────────────────────
df = pd.DataFrame(photos).sort_values("date")
outfile = Path(f"tagged_photos_{START_DATE}_to_{END_DATE}.xlsx")
df.to_excel(outfile, index=False)
print("Excel file written to", outfile.resolve())