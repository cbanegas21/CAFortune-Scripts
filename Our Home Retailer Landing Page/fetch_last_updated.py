#!/usr/bin/env python3
import json, re
from datetime import datetime
import pyodbc
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

SERVER = 'ca-data-server.database.windows.net'
DATABASE = 'CAFortuneDatabase'
USERNAME = 'sqladmin'
PASSWORD = 'Maxine2021.'
DRIVER   = '{ODBC Driver 17 for SQL Server}'
CONNECTION_TIMEOUT = 30
MAX_RETRIES = 5
RETRY_DELAY = 3
POWER_BI_ONLY = True

with open('assets/js/data.js','r', encoding='utf-8') as f:
    js = f.read()
reports = json.loads(js[js.index('=')+1:].strip().rstrip(';'))
reports = [r for r in reports if (not POWER_BI_ONLY) or r['type']=='powerbi']

def norm(s:str)->str:
    return re.sub(r'[^a-z0-9]', '', s.lower())

@retry(stop=stop_after_attempt(MAX_RETRIES),
       wait=wait_fixed(RETRY_DELAY),
       retry=retry_if_exception_type((pyodbc.OperationalError, pyodbc.InterfaceError)),
       reraise=True)
def connect():
    conn_str = (
        f"DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};"
        f"UID={USERNAME};PWD={PASSWORD};Connection Timeout={CONNECTION_TIMEOUT};"
        "Encrypt=yes;TrustServerCertificate=no;"
    )
    cn = pyodbc.connect(conn_str, autocommit=False)
    with cn.cursor() as cur: cur.execute("SELECT 1")
    return cn

def find_table_for_report(cur, report_name:str):
    m = re.search(r'(?:OUR HOME|Our Home)\s+(.*)', report_name)
    base = m.group(1) if m else report_name
    key = norm(base)
    cur.execute("SELECT name FROM sys.tables WHERE name LIKE 'Our Home %'")
    rows = [r[0] for r in cur.fetchall()]
    scored = []
    for t in rows:
        tn = norm(t.replace('Our Home','').strip())
        score = 0
        if key in tn or tn in key: score += 2
        toks1 = set(re.findall(r'[a-z0-9]+', key))
        toks2 = set(re.findall(r'[a-z0-9]+', tn))
        score += len(toks1 & toks2)
        if score>0: scored.append((score, t))
    if not scored: return None
    scored.sort(reverse=True)
    return scored[0][1]

def has_dateandtime(cur, schema:str, table:str)->bool:
    cur.execute("""
        SELECT 1
        FROM INFORMATION_SCHEMA.COLUMNS 
        WHERE TABLE_SCHEMA=? AND TABLE_NAME=? AND COLUMN_NAME='DateAndTime'
    """, (schema, table))
    return cur.fetchone() is not None

def fetch_max_date(cur, full_table_name:str):
    import re
    m = re.match(r'(?:(\w+)\.)?\[?([^.\]]+)\]?', full_table_name)
    if not m: return None
    schema = 'dbo' if m.group(1) is None else m.group(1)
    table  = m.group(2)
    if not has_dateandtime(cur, schema, table):
        return None
    q = f"SELECT TRY_CONVERT(datetime2, MAX([DateAndTime])) FROM [{schema}].[{table}]"
    cur.execute(q)
    r = cur.fetchone()
    return r[0]

def main():
    cn = connect()
    out = {}
    try:
        with cn.cursor() as cur:
            for r in reports:
                name = r['name']
                slug = r['slug']
                tbl = find_table_for_report(cur, name)
                if not tbl:
                    continue
                full_tbl = f'dbo.[{tbl}]' if not tbl.startswith('dbo.') else tbl
                dt = fetch_max_date(cur, full_tbl)
                if dt:
                    out[slug] = dt.strftime('%Y-%m-%d')
    finally:
        cn.close()
    with open('last_updated.json','w',encoding='utf-8') as f:
        json.dump(out, f, indent=2)
    print(f"Wrote last_updated.json with {len(out)} entries.")
if __name__ == '__main__':
    main()
