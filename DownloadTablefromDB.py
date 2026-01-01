# -*- coding: utf-8 -*-
"""
SELECT * FROM dbo.ExecutiveSummary_OurHome (con filtro de fechas) -> Excel (.xlsx)
- Lectura por lotes via pyodbc (sin límite de 50k)
- Escritura directa con XlsxWriter (sin pandas), multi-hoja automática
- Encabezado en cada hoja; rollover a Data1, Data2, ...

Ajusta la CONFIG para fechas, servidor y salida.
"""

import os
import logging
from contextlib import closing
from datetime import date, timedelta

import pyodbc
import xlsxwriter
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

# ==============================
# CONFIGURACIÓN
# ==============================

# Fechas (incluye el día final)
FROM_DATE = date(2025, 10, 31)
TO_DATE   = date(2025, 12, 6)

# Conexión Azure SQL
SERVER   = 'ca-data-server.database.windows.net'
DATABASE = 'CAFortuneDatabase'
USERNAME = os.getenv('SQL_USER', 'sqladmin')
PASSWORD = os.getenv('SQL_PASSWORD', 'Maxine2021.')
ODBC_DRIVER = '{ODBC Driver 18 for SQL Server}'  # o '{ODBC Driver 17 for SQL Server}'
CONNECTION_TIMEOUT = 30

# Exportación Excel
OUTPUT_XLSX = r'C:\Exports\ExecutiveSummary_2025-10-31_to_2025-12-06.xlsx'
CHUNK_SIZE = 50_000
MAX_ROWS_PER_SHEET = 1_048_000   # debajo del máximo 1,048,576 (incluye header)
SHEET_PREFIX = 'Data'            # Data1, Data2, ...

# Retries
MAX_RETRIES = 5
RETRY_DELAY = 10  # segundos

# Logging
logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)


# ==============================
# DB CONNECTION (con retry)
# ==============================

db_retry = retry(
    stop=stop_after_attempt(MAX_RETRIES),
    wait=wait_fixed(RETRY_DELAY),
    retry=retry_if_exception_type((pyodbc.OperationalError, pyodbc.InterfaceError)),
    reraise=True
)

@db_retry
def get_db_connection() -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={ODBC_DRIVER};"
        f"SERVER=tcp:{SERVER},1433;"
        f"DATABASE={DATABASE};"
        f"UID={USERNAME};PWD={PASSWORD};"
        f"Encrypt=yes;TrustServerCertificate=no;"
        f"Connection Timeout={CONNECTION_TIMEOUT};"
    )
    conn = pyodbc.connect(conn_str, autocommit=True)
    # ping
    with conn.cursor() as cur:
        cur.execute("SELECT 1")
        _ = cur.fetchone()
    return conn


# ==============================
# EXPORT LOGIC
# ==============================

SQL_SELECT_ALL = """
SELECT *
FROM dbo.ExecutiveSummary_OurHome
WHERE [Date] IS NOT NULL
  AND TRY_CONVERT(date, [Date]) >= ?
  AND TRY_CONVERT(date, [Date]) <  ?
ORDER BY TRY_CONVERT(datetime2, [Date]) ASC, TableName, Retailer;
"""

def export():
    os.makedirs(os.path.dirname(OUTPUT_XLSX) or ".", exist_ok=True)
    to_plus_one = TO_DATE + timedelta(days=1)

    total_rows = 0
    sheet_index = 1
    current_row = 0  # 0 = header row

    workbook = xlsxwriter.Workbook(OUTPUT_XLSX, {'constant_memory': True})
    header_fmt = workbook.add_format({'bold': True})
    worksheet = workbook.add_worksheet(f"{SHEET_PREFIX}{sheet_index}")

    with closing(get_db_connection()) as conn, conn.cursor() as cur:

        cur.arraysize = CHUNK_SIZE
        cur.execute(SQL_SELECT_ALL, (FROM_DATE, to_plus_one))

        # Obtener columnas a partir del cursor (SELECT *)
        col_names = [c[0] for c in cur.description]

        # Escribir header en la primera hoja
        worksheet.write_row(0, 0, col_names, header_fmt)
        current_row = 1

        while True:
            batch = cur.fetchmany(CHUNK_SIZE)
            if not batch:
                break

            for row in batch:
                # ¿necesitamos nueva hoja?
                if current_row >= MAX_ROWS_PER_SHEET:
                    sheet_index += 1
                    worksheet = workbook.add_worksheet(f"{SHEET_PREFIX}{sheet_index}")
                    worksheet.write_row(0, 0, col_names, header_fmt)
                    current_row = 1

                # Escribir fila (pyodbc.Row es iterable)
                worksheet.write_row(current_row, 0, list(row))
                current_row += 1
                total_rows += 1

    # Info sheet (opcional)
    info = workbook.add_worksheet("Info")
    info.write_row(0, 0, ["From", "To", "Rows", "Sheets"])
    info.write_row(1, 0, [str(FROM_DATE), str(TO_DATE), total_rows, sheet_index])

    workbook.close()
    logger.info(f"OK: {total_rows} filas exportadas a {OUTPUT_XLSX} en {sheet_index} hoja(s).")


# ==============================
# MAIN
# ==============================

if __name__ == "__main__":
    try:
        logger.info("Iniciando exportación (SELECT * -> Excel)…")
        export()
        logger.info("Exportación finalizada.")
    except Exception as e:
        logger.exception(f"ERROR en exportación: {e}")
        raise
