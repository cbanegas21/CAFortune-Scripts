"""
update_solely_heb_from_ui_facings_inventory.py

• Pick the Solely-HEB Excel export (UI).
• Match on timestamp and update 8 clean columns in dbo.solely_heb:
  mango_passion_fruit_facings (INT),  mango_passion_fruit_inventory (NVARCHAR),
  mango_guava_facings (INT),          mango_guava_inventory (NVARCHAR),
  mango_blueberry_facings (INT),      mango_blueberry_inventory (NVARCHAR),
  mango_strawberry_facings (INT),     mango_strawberry_inventory (NVARCHAR)

Notes:
- Inventory is TEXT and preserved as-is.
- Facings are numeric.
- Uses COALESCE so NULL inputs do NOT overwrite existing values.
"""

import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyodbc
from typing import Dict

# ───────────────────────────────────────────────────────────────────────────────
# 1) DB connection config  (FILL THESE)
# ───────────────────────────────────────────────────────────────────────────────
SERVER   = "ca-data-server.database.windows.net"
DATABASE = "CAFortuneDatabase"
UID      = "sqladmin"
PWD      = "Maxine2021."  # ← put the correct password here

TABLE_FQN = "dbo.solely_heb"
DB_DT_COL = "DateAndTime"

CONN_STR_18 = (
    f"DRIVER={{ODBC Driver 18 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={UID};PWD={PWD};"
    f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
)
CONN_STR_17 = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={SERVER};DATABASE={DATABASE};UID={UID};PWD={PWD};"
    f"Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
)

# ───────────────────────────────────────────────────────────────────────────────
# 2) Excel labels → DB columns (clean names with 'facings' and 'inventory')
# ───────────────────────────────────────────────────────────────────────────────
EXPECTED_TO_DB: Dict[str, str] = {
    # Mango & Passion Fruit
    "Fruit Gummies Mango & Passion Fruit  850023073907 | How many faces of product on shelf? (Solely SKU Audit)": "mango_passion_fruit_facings",
    "Fruit Gummies Mango & Passion Fruit  850023073907 | Inventory Levels (Solely SKU Audit)":                    "mango_passion_fruit_inventory",

    # Mango & Guava
    "Fruit Gummies Mango & Guava 85002307302 | How many faces of product on shelf? (Solely SKU Audit)":           "mango_guava_facings",
    "Fruit Gummies Mango & Guava 85002307302 | Inventory Levels (Solely SKU Audit)":                               "mango_guava_inventory",

    # Mango & Blueberry
    "Fruit Gummies Mango & Blueberry 85005299724  | How many faces of product on shelf? (Solely SKU Audit)":      "mango_blueberry_facings",
    "Fruit Gummies Mango & Blueberry 85005299724  | Inventory Levels (Solely SKU Audit)":                          "mango_blueberry_inventory",

    # Mango & Strawberry
    "Fruit Gummies Mango & Strawberry 85005299760  | How many faces of product on shelf? (Solely SKU Audit)":     "mango_strawberry_facings",
    "Fruit Gummies Mango & Strawberry 85005299760  | Inventory Levels (Solely SKU Audit)":                         "mango_strawberry_inventory",
}

# Desired SQL types per column: facings = INT, inventory = NVARCHAR(255)
def expected_sql_type(col_name: str) -> str:
    return "NVARCHAR(255)" if col_name.endswith("_inventory") else "INT"

EXCEL_DT_COL = "Date and time"  # Excel timestamp

# ───────────────────────────────────────────────────────────────────────────────
# Helpers
# ───────────────────────────────────────────────────────────────────────────────
def norm(txt: str) -> str:
    t = re.sub(r"\s+", " ", str(txt or "")).strip().lower()
    return t

def build_forgiving_map(actual_cols) -> Dict[str, str]:
    """Map EXPECTED labels to actual Excel column names (case/space-insensitive)."""
    actual_norm_map = {norm(c): c for c in actual_cols}
    resolved, missing = {}, []
    for expected in EXPECTED_TO_DB.keys():
        key = norm(expected)
        hit = actual_norm_map.get(key)
        if hit:
            resolved[expected] = hit
        else:
            missing.append(expected)
    if missing:
        print("WARNING: Expected columns not found exactly; creating blanks for:")
        for m in missing:
            print("  -", m)
    return resolved

def connect_db() -> pyodbc.Connection:
    try:
        return pyodbc.connect(CONN_STR_18, autocommit=False)
    except pyodbc.Error:
        print("Driver 18 failed, trying Driver 17…")
        return pyodbc.connect(CONN_STR_17, autocommit=False)

def parse_excel_datetime(series: pd.Series) -> pd.Series:
    """Normalize odd spaces & parse with explicit formats; fallback to dateutil."""
    s = series.astype(str)
    s = (
        s.str.replace("\u202f", " ", regex=False)  # narrow no-break space
         .str.replace("\xa0", " ", regex=False)    # no-break space
         .str.strip()
    )
    fmts = [
        "%m/%d/%Y %I:%M:%S %p",  # e.g., 8/1/2025 1:42:01 PM
        "%m/%d/%Y %H:%M:%S",     # 24h
        "%d/%m/%Y %I:%M:%S %p",  # day-first
        "%d/%m/%Y %H:%M:%S",
    ]
    dt = None
    for fmt in fmts:
        dt_try = pd.to_datetime(s, format=fmt, errors="coerce")
        if dt_try.notna().mean() >= 0.85:
            dt = dt_try; break
    if dt is None:
        dt = pd.to_datetime(s, errors="coerce")
    return dt

def ensure_column_type(cur: pyodbc.Cursor, table_fqn: str, col: str, sql_type: str):
    """
    Ensure column exists with the desired SQL type.
    If missing -> ADD.
    If type differs -> ALTER COLUMN (NULL).
    """
    cur.execute("""
SELECT t.name
FROM sys.columns c
JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID(?) AND c.name = ?;
""", table_fqn, col)
    row = cur.fetchone()
    if row is None:
        cur.execute(f"EXEC('ALTER TABLE {table_fqn} ADD {col} {sql_type} NULL;')")
    else:
        current_type = row[0].upper()
        # NVARCHAR(max) etc. normalize comparison a bit
        if sql_type.startswith("NVARCHAR") and not current_type.startswith("NVARCHAR"):
            cur.execute(f"EXEC('ALTER TABLE {table_fqn} ALTER COLUMN {col} {sql_type} NULL;')")
        elif sql_type == "INT" and current_type != "INT":
            cur.execute(f"EXEC('ALTER TABLE {table_fqn} ALTER COLUMN {col} INT NULL;')")

# ───────────────────────────────────────────────────────────────────────────────
# Main
# ───────────────────────────────────────────────────────────────────────────────
def main():
    # 1) Pick Excel
    root = tk.Tk(); root.withdraw()
    xlsx_path = filedialog.askopenfilename(
        title="Select Solely-HEB Excel export",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not xlsx_path:
        messagebox.showerror("No file selected", "You must choose an Excel file.")
        return

    # 2) Load Excel
    df = pd.read_excel(xlsx_path, dtype=str)
    if EXCEL_DT_COL not in df.columns:
        messagebox.showerror("Column not found", f"'{EXCEL_DT_COL}' not in the Excel.")
        return

    resolved_map = build_forgiving_map(df.columns)

    # 3) Parse timestamp -> seconds precision
    dt = parse_excel_datetime(df[EXCEL_DT_COL])
    # If Excel is local time but SQL is UTC, you can do:
    # dt = dt.dt.tz_localize("America/Tegucigalpa").dt.tz_convert("UTC").tz_localize(None)

    df["dt_key"] = dt.dt.floor("S")
    df = df.dropna(subset=["dt_key"])

    # 4) Working frame with mapped columns
    work_cols = ["dt_key"]
    for expected_label in EXPECTED_TO_DB.keys():
        actual_col = resolved_map.get(expected_label)
        if actual_col and actual_col in df.columns:
            work_cols.append(actual_col)
        else:
            df[expected_label] = pd.NA
            work_cols.append(expected_label)
    wdf = df[work_cols].copy()

    # 5) Connect and ensure column types (INT for facings, NVARCHAR for inventory)
    try:
        cn = connect_db()
    except pyodbc.Error as e:
        messagebox.showerror(
            "DB Connection Failed",
            "Could not connect to SQL Server.\n\n"
            "• Verify SERVER/DATABASE/UID/PWD.\n"
            "• Try the same credentials in SSMS.\n"
            "• Ensure ODBC Driver 18/17 is installed.\n"
            "• Ensure your IP is allowed in Azure SQL firewall."
        )
        raise

    cur = cn.cursor()
    for expected_label, db_col in EXPECTED_TO_DB.items():
        ensure_column_type(cur, TABLE_FQN, db_col, expected_sql_type(db_col))
    cn.commit()

    # 6) Build batched parameters
    ordered_expected = list(EXPECTED_TO_DB.keys())
    set_cols = [EXPECTED_TO_DB[k] for k in ordered_expected]

    # Use COALESCE so NULL params don't wipe existing data
    update_sql = f"""
UPDATE {TABLE_FQN}
SET
  {set_cols[0]} = COALESCE(?, {set_cols[0]}),
  {set_cols[1]} = COALESCE(?, {set_cols[1]}),
  {set_cols[2]} = COALESCE(?, {set_cols[2]}),
  {set_cols[3]} = COALESCE(?, {set_cols[3]}),
  {set_cols[4]} = COALESCE(?, {set_cols[4]}),
  {set_cols[5]} = COALESCE(?, {set_cols[5]}),
  {set_cols[6]} = COALESCE(?, {set_cols[6]}),
  {set_cols[7]} = COALESCE(?, {set_cols[7]})
WHERE CAST({DB_DT_COL} AS datetime2(0)) = CAST(? AS datetime2(0));
"""

    rows = []
    facings_cols = {c for c in set_cols if c.endswith("_facings")}
    inventory_cols = {c for c in set_cols if c.endswith("_inventory")}

    for _, r in wdf.iterrows():
        param_values = []
        for expected_label in ordered_expected:
            db_col = EXPECTED_TO_DB[expected_label]
            actual_col = resolved_map.get(expected_label, expected_label)
            raw_val = r.get(actual_col)

            if db_col in facings_cols:
                # numeric coercion for facings
                try:
                    v = None if pd.isna(raw_val) else int(float(str(raw_val).strip()))
                except Exception:
                    v = None
                param_values.append(v)
            else:
                # inventory: keep as clean text
                if pd.isna(raw_val):
                    param_values.append(None)
                else:
                    s = str(raw_val).strip()
                    param_values.append(s if s != "" else None)

        # datetime key formatted to seconds
        param_values.append(pd.Timestamp(r["dt_key"]).strftime("%Y-%m-%d %H:%M:%S"))
        rows.append(tuple(param_values))

    print(f"Prepared {len(rows)} row updates.")

    # 7) Execute
    cur.fast_executemany = True
    cur.executemany(update_sql, rows)
    cn.commit()

    cur.close(); cn.close()
    messagebox.showinfo("Done", f"{len(rows):,} rows updated in {TABLE_FQN}.")

if __name__ == "__main__":
    main()
