import requests
import pyodbc
import re
from requests.auth import HTTPBasicAuth
import time
from datetime import datetime, timedelta

# Credenciales para conectarse a la base de datos
conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=ca-data-server.database.windows.net;DATABASE=CAFortuneDatabase;UID=sqladmin;PWD=Maxine2021.'

# Credenciales de la API
username = '80941603-F785-4E0F-8AB1-ED798E54F88C'
password = '17BCD0FD-94DD-4B0C-B059-76D68C1145A8'

# Headers de la API
headers = {
    'Content-Type': 'application/json'
}

# Parámetros para el control de reintentos
max_retries = 5
retry_delay = 5  # Segundos de espera entre reintentos

# Caché local para columnas y tablas verificadas
verified_tables = {}

# Función para conectarse a la base de datos
def connect_to_db():
    try:
        return pyodbc.connect(conn_str)
    except pyodbc.Error:
        return None

# Función para extraer datos de la API
def fetch_data_from_api(last_form_id):
    url = f"https://api.repsly.com/v3/export/forms/{last_form_id}"
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))
            if response.status_code == 200:
                return response.json()
            else:
                raise Exception(f"HTTP error {response.status_code}")
        except Exception:
            if attempt + 1 == max_retries:
                return None
            time.sleep(retry_delay)

# Función para limpiar nombres de columnas y tablas
def escape_name(name, max_length=128):
    cleaned_name = re.sub(r'[^\w]', '_', name)
    cleaned_name = cleaned_name[:max_length]
    cleaned_name = cleaned_name.strip('_')
    return cleaned_name

# Verificar si la tabla ya existe (sin llaves)
def table_exists(cursor, table_name):
    if table_name in verified_tables:
        return True
    
    query = f"""
        SELECT COUNT(*)
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_NAME = '{table_name}'
    """
    cursor.execute(query)
    exists = cursor.fetchone()[0] > 0
    if exists:
        verified_tables[table_name] = set()  # Agregar la tabla al diccionario
    return exists

# Crear una tabla si no existe, asegurando que los nombres de columnas sean únicos y normalizados
def create_table(cursor, table_name, columns):
    cleaned_columns = []
    unique_columns = set()  # Conjunto para evitar columnas duplicadas

    for column in columns:
        cleaned_column = escape_name(column.strip().lower())
        if len(cleaned_column) <= 128 and cleaned_column not in unique_columns:
            cleaned_columns.append(f"[{cleaned_column}] NVARCHAR(MAX)")
            unique_columns.add(cleaned_column)  # Agregar la columna al conjunto
        else:
            print(f"Columna duplicada o inválida detectada y omitida: {cleaned_column}")
    
    if cleaned_columns:
        create_sql = f"CREATE TABLE [{table_name}] ({', '.join(cleaned_columns)})"
        cursor.execute(create_sql)
        verified_tables[table_name] = set(cleaned_columns)  # Actualizar caché

# Verificar si la columna ya existe en la tabla (usando el nombre sin llaves)
def column_exists(cursor, table_name, column_name):
    if column_name in verified_tables.get(table_name, set()):
        return True
    
    query = f"""
        SELECT COUNT(*)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = '{table_name[1:-1]}' AND COLUMN_NAME = '{column_name}'
    """
    cursor.execute(query)
    exists = cursor.fetchone()[0] > 0
    if exists:
        verified_tables[table_name] = verified_tables.get(table_name, set())  # Crear entrada en el caché de columnas
        verified_tables[table_name].add(column_name)
    return exists

# Función para convertir la fecha en formato "/Date(...)/" a datetime
def convert_api_date(api_date_string):
    # Verificar si el formato es tipo "/Date(....)/"
    if api_date_string.startswith("/Date(") and api_date_string.endswith(")/"):
        timestamp = int(api_date_string[6:19]) // 1000  # Convertir a segundos
        return datetime(1970, 1, 1) + timedelta(seconds=timestamp)
    else:
        # Si el formato no es compatible, devolver None
        return None

# Crear columnas faltantes en una tabla de manera optimizada
def add_missing_columns(cursor, table_name, row_data):
    for column_name, value in row_data.items():
        cleaned_column_name = escape_name(column_name)
        if len(cleaned_column_name) > 128:
            cleaned_column_name = cleaned_column_name[:128]

        # Verificar y añadir columna si es necesario
        if not column_exists(cursor, table_name, cleaned_column_name):
            alter_sql = f"ALTER TABLE [{escape_name(table_name)}] ADD [{cleaned_column_name}] NVARCHAR(MAX)"
            try:
                cursor.execute(alter_sql)
                print(f"Columna añadida: {cleaned_column_name} en tabla {table_name}")
            except pyodbc.Error as e:
                print(f"Error al añadir la columna {cleaned_column_name}: {str(e)}")

def insert_data(cursor, table_name, row_data):
    table_name = f"[{escape_name(table_name)}]"
    
    try:
        escaped_columns = [escape_name(col) for col in row_data.keys() if len(escape_name(col)) <= 128]
        if not escaped_columns:
            return

        columns = ', '.join([f"[{col}]" for col in escaped_columns])
        placeholders = ', '.join(['?' for _ in row_data.values()])
        insert_sql = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"

        cursor.execute(insert_sql, tuple(row_data.values()))
        print(f"Datos insertados en la tabla {table_name}")

    except pyodbc.Error as e:
        error_message = str(e)

        # Si el error es por una columna no válida (columna nueva no encontrada), la agregamos
        if 'Invalid column name' in error_message:
            start_index = error_message.find("Invalid column name '") + len("Invalid column name '")
            end_index = error_message.find("'", start_index)
            invalid_column = error_message[start_index:end_index]
            print(f"Columna nueva detectada: {invalid_column}. Agregando la columna a la tabla {table_name}.")

            # Agregar la nueva columna a la tabla
            try:
                add_column_sql = f"ALTER TABLE {table_name} ADD [{invalid_column}] NVARCHAR(MAX)"
                cursor.execute(add_column_sql)
                print(f"Columna {invalid_column} añadida a la tabla {table_name}. Reintentando inserción.")

                # Reintentar la inserción después de agregar la columna
                insert_data(cursor, table_name, row_data)
            except pyodbc.Error as inner_error:
                print(f"Error al agregar la columna {invalid_column}: {inner_error}")
        elif 'Column name' in error_message and 'specified more than once' in error_message:
            # Si se detecta un error de columna duplicada, lo manejamos y continuamos
            print(f"Columna duplicada detectada: {error_message}")
        else:
            # Si el error no es por columna inválida o duplicada, lo reportamos
            print(f"Error insertando en la tabla {table_name}: {e}")

def process_form(cursor, form_name, form_data):
    table_name = escape_name(form_name)
    row_data = {
        "FormID": form_data.get("FormID"),
        "FormName": form_data.get("FormName"),
        "ClientCode": form_data.get("ClientCode"),
        "ClientName": form_data.get("ClientName"),
        "DateAndTime": form_data.get("DateAndTime"),
        "RepresentativeCode": form_data.get("RepresentativeCode"),
        "RepresentativeName": form_data.get("RepresentativeName"),
        "StreetAddress": form_data.get("StreetAddress"),
        "ZIP": form_data.get("ZIP"),
        "ZIPExt": form_data.get("ZIPExt"),
        "City": form_data.get("City"),
        "State": form_data.get("State"),
        "Country": form_data.get("Country"),
        "Email": form_data.get("Email"),
        "Phone": form_data.get("Phone"),
        "Mobile": form_data.get("Mobile"),
        "Territory": form_data.get("Territory"),
        "Longitude": form_data.get("Longitude"),
        "Latitude": form_data.get("Latitude"),
        "SignatureURL": form_data.get("SignatureURL"),
        "VisitStart": form_data.get("VisitStart"),
        "VisitEnd": form_data.get("VisitEnd")
    }

    for item in form_data.get("Items", []):
        field_name = item.get("Field")
        if field_name:
            cleaned_field_name = escape_name(field_name).lower().replace(" ", "_")
            if len(cleaned_field_name) <= 128:
                if cleaned_field_name in row_data:
                    # Ignorar la columna duplicada
                    continue
                else:
                    row_data[cleaned_field_name] = item.get("Value")

    # Verificar si la tabla ya existe o crearla
    if not table_exists(cursor, table_name):
        create_table(cursor, table_name, row_data.keys())
    
    insert_data(cursor, table_name, row_data)

# Función para insertar un registro en ApiUpdateLogs
def insert_log(cursor, last_form_id, total_forms_inserted, last_form_date):
    # Convertir la fecha de la API si es necesario
    converted_date = convert_api_date(last_form_date)
    
    # Si no es una fecha válida, usar la fecha actual
    if not converted_date:
        print(f"Formato de fecha inválido detectado: {last_form_date}. Usando la fecha actual.")
        converted_date = datetime.now()

    # Insertar el log con la fecha convertida
    insert_log_sql = """
    INSERT INTO ApiUpdateLogs (LastFormID, TotalFormsInserted, LastFormDate)
    VALUES (?, ?, ?)
    """
    cursor.execute(insert_log_sql, (last_form_id, total_forms_inserted, converted_date.isoformat()))
    print(f"Log insertado para el LastFormID: {last_form_id} con {total_forms_inserted} formularios y fecha: {converted_date}.")

# Procesar datos y continuar con el siguiente Last Form ID
# Procesar datos y continuar con el siguiente Last Form ID
def process_data(starting_last_form_id):
    conn = connect_to_db()
    if conn is None:
        return

    cursor = conn.cursor()
    last_form_id = starting_last_form_id
    forms_per_commit = 50  # Número de formularios por transacción
    total_forms_inserted = 0  # Contador para total de formularios procesados
    last_form_date = None  # Inicializa la variable para la fecha del último formulario

    while last_form_id is not None:
        forms_data = fetch_data_from_api(last_form_id)
        if forms_data is None or len(forms_data.get("Forms", [])) == 0:
            break

        form_count = 0
        print(f"Procesando Last Form ID: {last_form_id}")

        for form in forms_data.get("Forms", []):
            form_name = form.get("FormName", "UnknownForm")
            print(f"Preparando formulario: {form_name} con FormID: {form.get('FormID')}")

            process_form(cursor, form_name, form)
            form_count += 1
            total_forms_inserted += 1
            last_form_date = form.get("DateAndTime")  # Guarda la fecha del último form procesado

            # Hacer commit cada 'forms_per_commit' formularios
            if form_count % forms_per_commit == 0:
                conn.commit()
                print(f"Los {forms_per_commit} formularios han sido subidos a la base de datos.")

        # Hacer commit final para formularios restantes si hay menos de 50
        if form_count % forms_per_commit != 0:
            conn.commit()
            print(f"Todos los formularios restantes han sido subidos a la base de datos.")

        # Inserta el log con el último LastFormID, total de forms y fecha del último form
        if total_forms_inserted > 0 and last_form_date:
            insert_log(cursor, last_form_id, total_forms_inserted, last_form_date)

        last_form_id = forms_data['MetaCollectionResult'].get('LastID', None)
        print(f"Todos los formularios para Last Form ID {last_form_id} han sido subidos.")

    cursor.close()
    conn.close()


# Ejecutar el proceso
starting_last_form_id = "222381154"
process_data(starting_last_form_id)
