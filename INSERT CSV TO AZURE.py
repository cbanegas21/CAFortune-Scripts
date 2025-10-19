import os, logging
import pandas as pd
import pyodbc
from contextlib import closing

# ── CONFIG ──────────────────────────────────────────────────────────────
SERVER   = 'ca-data-server.database.windows.net'
DATABASE = 'CAFortuneDatabase'
USERNAME = 'sqladmin'
PASSWORD = 'Maxine2021.'
DRIVER   = '{ODBC Driver 17 for SQL Server}'

EXCEL_FILE   = 'The fresh market fix.xlsx'         # <──  nombre exacto del archivo
TARGET_TABLE = 'pur_gum_the_fresh_market'  # <──  nombre de la tabla destino
CHUNK_SIZE   = 1000
# ────────────────────────────────────────────────────────────────────────

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s  %(levelname)s  %(message)s')
log = logging.getLogger(__name__)

# ── DB HELPERS ──────────────────────────────────────────────────────────
def get_conn():
    return pyodbc.connect(
        f'DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};'
        f'UID={USERNAME};PWD={PASSWORD};Encrypt=yes;TrustServerCertificate=no;',
        autocommit=False
    )

def sql_table_exists(cur, tbl):
    cur.execute("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME=?", tbl)
    return cur.fetchone() is not None

def create_table(cur, tbl, cols):
    col_defs = ", ".join(f"[{c}] NVARCHAR(MAX)" for c in cols)
    cur.execute(f"CREATE TABLE [{tbl}] ({col_defs})")
    log.info("Tabla %s creada con %d columnas", tbl, len(cols))

def get_sql_types(cur, tbl):
    rows = cur.execute(
        "SELECT COLUMN_NAME, DATA_TYPE "
        "FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?", tbl
    ).fetchall()
    return {r.COLUMN_NAME: r.DATA_TYPE.lower() for r in rows}

def cast_df(df, sql_types):
    for col_sql, dtype in sql_types.items():
        if col_sql not in df.columns:
            df[col_sql] = None                     # columna falta en Excel
        elif dtype in {"bigint","int","smallint","decimal","numeric","float"}:
            df[col_sql] = pd.to_numeric(df[col_sql], errors="coerce")
        elif dtype.startswith("date") or dtype.startswith("time"):
            df[col_sql] = pd.to_datetime(df[col_sql], errors="coerce")
    # quita columnas que no existen en SQL
    return df[[c for c in df.columns if c in sql_types]]

def insert_chunks(cur, tbl, df):
    cols = list(df.columns)
    ph   = ", ".join("?" * len(cols))
    sql  = f"INSERT INTO [{tbl}] ({', '.join(f'[{c}]' for c in cols)}) VALUES ({ph})"
    df   = df.where(pd.notnull(df), None)

    for start in range(0, len(df), CHUNK_SIZE):
        chunk = df.iloc[start:start+CHUNK_SIZE]
        cur.fast_executemany = True
        cur.executemany(sql, chunk.values.tolist())
        log.info("Insertados %d filas (%d-%d)", len(chunk), start+1, start+len(chunk))

# ── MAIN ────────────────────────────────────────────────────────────────
def main():
    here = os.path.abspath(os.path.dirname(__file__))
    xlsx_path = os.path.join(here, EXCEL_FILE)
    if not os.path.exists(xlsx_path):
        log.error("No se encontró %s", xlsx_path)
        return

    # lee la primera hoja, todo como texto
    df = pd.read_excel(xlsx_path, dtype=str, engine="openpyxl")
    df.columns = [c.strip().replace(" ", "_") for c in df.columns]

    with closing(get_conn()) as conn:
        cur = conn.cursor()

        if not sql_table_exists(cur, TARGET_TABLE):
            create_table(cur, TARGET_TABLE, df.columns)

        sql_types = get_sql_types(cur, TARGET_TABLE)
        df        = cast_df(df, sql_types)
        insert_chunks(cur, TARGET_TABLE, df)
        conn.commit()

    log.info("✔  Carga completa: %d filas añadidas a %s", len(df), TARGET_TABLE)

if __name__ == "__main__":
    main()
