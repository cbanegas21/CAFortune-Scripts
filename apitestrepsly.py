"""
Purchase‑orders pull from Repsly API  ‑‑ 2025‑07‑23

• Starts at lastDocumentID = 0  (change to your saved LastID for incremental sync)
• Handles 50‑row paging until MetaCollectionResult.TotalCount == 0
• Prints:  PurchaseOrderID • DocumentNo • ClientName • Date • <item‑count> items
• Keeps full JSON in all_orders  (list of dicts)
"""

import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
import sys, time, json

# ── Console UTF‑8 fix for Windows (prevents BOM \ufeff errors) ────────────────
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ── Credentials ───────────────────────────────────────────────────────────────
USERNAME = "80941603-F785-4E0F-8AB1-ED798E54F88C"
PASSWORD = "17BCD0FD-94DD-4B0C-B059-76D68C1145A8"

# ── Endpoint base ─────────────────────────────────────────────────────────────
BASE_URL = "https://api.repsly.com/v3/export/purchaseorders"

HEADERS = {"Accept": "application/json"}

# ── Paging setup ──────────────────────────────────────────────────────────────
last_document_id = 0      # ←  set to saved LastID for incremental runs
all_orders       = []

MAX_RETRIES = 5
RETRY_WAIT  = 5  # seconds


def ms_date_to_str(ms_date: str) -> str:
    """Convert Repsly /Date(1698045694000+0000)/ → 'YYYY‑mm‑dd HH:MM' (UTC)."""
    try:
        ts = int(ms_date[6:19]) / 1000.0
        return datetime.utcfromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return ms_date


def count_items(po: dict) -> int:
    """
    Items are flattened:  Item\<LineNo>\LineNo, Item\<LineNo>\ProductCode, …
    Count unique <LineNo> values appearing right after the first backslash.
    """
    line_numbers = {
        part.split("\\")[1]
        for part in po.keys()
        if part.startswith("Item\\") and part.endswith("\\LineNo")
    }
    return len(line_numbers)


while True:
    url = f"{BASE_URL}/{last_document_id}"

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url,
                auth=HTTPBasicAuth(USERNAME, PASSWORD),
                headers=HEADERS,
                timeout=60,
            )
            if resp.status_code != 200:
                raise RuntimeError(
                    f"HTTP {resp.status_code} :: {resp.text[:200]}…"
                )

            payload = resp.json()

            # Array key can vary slightly
            batch = (
                payload.get("PurchaseOrders")
                or payload.get("purchaseOrders")
                or payload.get("Orders")
                or []
            )

            meta         = payload.get("MetaCollectionResult") or {}
            total_count  = meta.get("TotalCount", 0)
            last_id      = meta.get("LastID", 0)

            if not batch and total_count == 0:
                print(f"\nNo more records (LastID={last_document_id}). Done.")
                print(f"TOTAL PURCHASE ORDERS RETRIEVED: {len(all_orders)}")
                # ── Optional: write to JSON file ───────────────────────────
                # with open("purchase_orders.json", "w", encoding="utf-8") as f:
                #     json.dump(all_orders, f, ensure_ascii=False, indent=2)
                sys.exit(0)

            # ── Console summary ───────────────────────────────────────────
            for po in batch:
                num_items = count_items(po)
                print(
                    f"{po['PurchaseOrderID']} • {po['DocumentNo']} • "
                    f"{po['ClientName']} • {ms_date_to_str(po['DateAndTime'])} • "
                    f"{num_items} items"
                )

            all_orders.extend(batch)
            last_document_id = last_id        # prepare next loop
            break  # success  →  leave retry loop

        except Exception as e:
            print(f"Attempt {attempt}/{MAX_RETRIES} failed: {e}")
            if attempt == MAX_RETRIES:
                sys.exit("Max retries reached; aborting.")
            time.sleep(RETRY_WAIT)
