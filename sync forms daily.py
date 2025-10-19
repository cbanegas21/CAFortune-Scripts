import os
import json
import time
import logging
import tempfile
import requests
import pandas as pd
import pyodbc
from datetime import datetime
from requests.auth import HTTPBasicAuth
from contextlib import closing
from tenacity import retry, stop_after_attempt, wait_fixed, retry_if_exception_type

# Configuración                                             
CHUNK_SIZE = 1000
MAX_COL_NAME_LEN = 100
MAX_RETRIES = 5
API_TIMEOUT = 30
RETRY_DELAY = 15  # segundos
CONNECTION_TIMEOUT = 30  # segundos

# Configuración de Azure SQL
SERVER = 'ca-data-server.database.windows.net'
DATABASE = 'CAFortuneDatabase'
USERNAME = 'sqladmin'
PASSWORD = 'Maxine2021.'
DRIVER = '{ODBC Driver 17 for SQL Server}'

# Configuración de API
API_USERNAME = '80941603-F785-4E0F-8AB1-ED798E54F88C'
API_PASSWORD = '17BCD0FD-94DD-4B0C-B059-76D68C1145A8'
API_URL = 'https://api.repsly.com/v3/export/forms/'


# Configurar logging
logging.basicConfig(
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Decorador para reintentos de conexión
db_retry = retry(
    stop=stop_after_attempt(MAX_RETRIES),
    wait=wait_fixed(RETRY_DELAY),
    retry=retry_if_exception_type((pyodbc.OperationalError, pyodbc.InterfaceError)),
    before_sleep=lambda _: logger.warning("Reintentando conexión a la base de datos..."),
    reraise=True
)

@db_retry
def get_db_connection():
    conn_str = (
        f'DRIVER={DRIVER};'
        f'SERVER={SERVER};'
        f'DATABASE={DATABASE};'
        f'UID={USERNAME};'
        f'PWD={PASSWORD};'
        f'Connection Timeout={CONNECTION_TIMEOUT};'
        'Encrypt=yes;'
        'TrustServerCertificate=no;'
    )
    conn = pyodbc.connect(conn_str, autocommit=False)
    
    # Verificar conexión activa
    with conn.cursor() as cursor:
        cursor.execute("SELECT 1")
    return conn

def clean_column(field_name):
    base = field_name.strip().replace(" ", "_").replace("/", "_").replace("-", "_")
    cleaned = ''.join(c for c in base if c.isalnum() or c == "_").lower()
    return cleaned[:MAX_COL_NAME_LEN]

def convert_date(ms_date):
    if not ms_date or not ms_date.startswith("/Date("):
        return None
    try:
        timestamp = int(ms_date[6:19])
        return datetime.utcfromtimestamp(timestamp / 1000)
    except Exception as e:
        logger.error(f"Error converting date: {str(e)}")
        return None

def ensure_table_exists(cursor, table_name, column_order):
    try:
        # Verificar si la tabla existe
        cursor.execute("SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = ?", table_name)
        table_exists = cursor.fetchone()
        
        if not table_exists:
            # Crear nueva tabla
            columns_sql = []
            seen_columns = set()
            for col in column_order:
                col_lower = col.lower()
                if col_lower in seen_columns:
                    continue
                seen_columns.add(col_lower)
                col_type = "DATETIME2" if col_lower in {"dateandtime", "visitstart", "visitend"} else "NVARCHAR(MAX)"
                columns_sql.append(f"[{col}] {col_type}")
            
            create_sql = f"CREATE TABLE [{table_name}] ({', '.join(columns_sql)})"
            cursor.execute(create_sql)
            logger.info(f"Tabla {table_name} creada exitosamente")
        else:
            # Agregar columnas faltantes
            cursor.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?", table_name)
            existing_columns = {row.COLUMN_NAME.lower() for row in cursor.fetchall()}
            
            for col in column_order:
                col_lower = col.lower()
                if col_lower not in existing_columns:
                    col_type = "DATETIME2" if col_lower in {"dateandtime", "visitstart", "visitend"} else "NVARCHAR(MAX)"
                    cursor.execute(f"ALTER TABLE [{table_name}] ADD [{col}] {col_type}")
                    logger.info(f"Columna {col} agregada a {table_name}")
        
        cursor.commit()
    except pyodbc.Error as e:
        cursor.rollback()
        logger.error(f"Error creando/actualizando tabla {table_name}: {str(e)}")
        raise

def insert_chunk(cursor, table_name, df, column_order):
    placeholders = ', '.join(['?'] * len(column_order))
    columns = ', '.join(f'[{col}]' for col in column_order)
    insert_sql = f"INSERT INTO [{table_name}] ({columns}) VALUES ({placeholders})"
    
    # Preparar datos
    safe_df = df[column_order].copy()
    for col in safe_df.columns:
        if pd.api.types.is_datetime64_any_dtype(safe_df[col]):
            safe_df[col] = safe_df[col].dt.tz_localize(None)
        safe_df[col] = safe_df[col].astype(object).where(pd.notnull(safe_df[col]), None)
    
    # Insertar en chunks
    total_rows = len(safe_df)
    for i in range(0, total_rows, CHUNK_SIZE):
        chunk = safe_df.iloc[i:i+CHUNK_SIZE]
        try:
            cursor.executemany(insert_sql, chunk.values.tolist())
            cursor.commit()
            logger.info(f"Insertados {len(chunk)} registros en {table_name} (lote {i//CHUNK_SIZE + 1})")
        except pyodbc.Error as e:
            cursor.rollback()
            logger.error(f"Error insertando datos en {table_name}: {str(e)}")
            raise

def get_last_synced_form_id(cursor):
    try:
        cursor.execute("SELECT TOP 1 last_form_id FROM sync_log ORDER BY last_sync_time DESC")
        row = cursor.fetchone()
        return row[0] if row else 237881055
    except pyodbc.Error as e:
        logger.error(f"Error obteniendo último ID sincronizado: {str(e)}")
        return 237881055

def update_sync_log(cursor, last_form_id):
    try:
        cursor.execute("INSERT INTO sync_log (last_form_id, last_sync_time) VALUES (?, GETDATE())", last_form_id)
        cursor.commit()
        logger.info(f"Log de sincronización actualizado con ID: {last_form_id}")
    except pyodbc.Error as e:
        cursor.rollback()
        logger.error(f"Error actualizando log de sincronización: {str(e)}")
        raise

def process_entry(entry):
    row = {
        "formid": entry.get("FormID"),
        "clientcode": entry.get("ClientCode"),
        "clientname": entry.get("ClientName"),
        "dateandtime": convert_date(entry.get("DateAndTime")),
        "representativecode": entry.get("RepresentativeCode"),
        "representativename": entry.get("RepresentativeName"),
        "streetaddress": entry.get("StreetAddress"),
        "zip": entry.get("ZIP"),
        "city": entry.get("City"),
        "state": entry.get("State"),
        "country": entry.get("Country"),
        "email": entry.get("Email"),
        "phone": entry.get("Phone"),
        "mobile": entry.get("Mobile"),
        "territory": entry.get("Territory"),
        "longitude": entry.get("Longitude"),
        "latitude": entry.get("Latitude"),
        "signatureurl": entry.get("SignatureURL"),
        "visitstart": convert_date(entry.get("VisitStart")),
        "visitend": convert_date(entry.get("VisitEnd")),
        "visitid": entry.get("VisitID"),
    }
    
    seen = set(row.keys())
    for item in entry.get("Items", []):
        col = clean_column(item["Field"])
        if col and col not in seen and len(col) <= MAX_COL_NAME_LEN:
            row[col] = item.get("Value")
            seen.add(col)
    
    return clean_column(entry["FormName"]), row

@retry(
    stop=stop_after_attempt(3),
    wait=wait_fixed(10),
    retry=retry_if_exception_type(requests.RequestException),
    reraise=True
)
def fetch_new_forms():
    try:
        with closing(get_db_connection()) as conn:
            with conn.cursor() as cursor:
                last_form_id = get_last_synced_form_id(cursor)
        
        all_forms = []
        current_last_id = last_form_id
        
        while True:
            try:
                response = requests.get(
                    f"{API_URL}{current_last_id}",
                    auth=HTTPBasicAuth(API_USERNAME, API_PASSWORD),
                    timeout=API_TIMEOUT
                )
                response.raise_for_status()
                
                data = response.json()
                forms = data.get("Forms", [])
                
                if not forms:
                    break
                    
                all_forms.extend(forms)
                new_last_id = data["MetaCollectionResult"]["LastID"]
                
                if new_last_id <= current_last_id:
                    break
                    
                current_last_id = new_last_id
                logger.info(f"Obtenidos {len(forms)} formularios, último ID: {current_last_id}")
                
            except (requests.RequestException, json.JSONDecodeError) as e:
                logger.error(f"API Error: {str(e)}")
                break
        
        return all_forms, current_last_id
    except Exception as e:
        logger.error(f"Error general obteniendo formularios: {str(e)}")
        raise

def main():
    logger.info("Iniciando proceso de sincronización")
    try:
        # Obtener datos de la API
        forms, new_last_id = fetch_new_forms()
        if not forms:
            logger.info("No hay nuevos formularios para procesar")
            return
        
        # Procesar datos
        form_tables = {}
        for entry in forms:
            form_name, row = process_entry(entry)
            if form_name not in form_tables:
                form_tables[form_name] = []
            form_tables[form_name].append(row)
        
        # Insertar en base de datos
        attempt = 1
        while attempt <= MAX_RETRIES:
            try:
                with closing(get_db_connection()) as conn:
                    conn.autocommit = False
                    cursor = conn.cursor()
                    cursor.fast_executemany = True
                    
                    for form_name, records in form_tables.items():
                        df = pd.DataFrame(records)
                        column_order = [col for col in df.columns if col]
                        
                        ensure_table_exists(cursor, form_name, column_order)
                        insert_chunk(cursor, form_name, df, column_order)
                        logger.info(f"Insertados {len(df)} registros en {form_name}")
                    
                    update_sync_log(cursor, new_last_id)
                    logger.info("Sincronización completada exitosamente")
                    break
                    
            except pyodbc.OperationalError as e:
                logger.error(f"Intento {attempt} fallido: {str(e)}")
                if attempt == MAX_RETRIES:
                    raise
                logger.info(f"Reintentando en {RETRY_DELAY} segundos...")
                time.sleep(RETRY_DELAY)
                attempt += 1
                
            except Exception as e:
                logger.error(f"Error durante operaciones de base de datos: {str(e)}")
                raise
                
    except Exception as e:
        logger.error(f"Error crítico: {str(e)}")
    finally:
        logger.info("Proceso de sincronización finalizado")

if __name__ == '__main__':
    main()