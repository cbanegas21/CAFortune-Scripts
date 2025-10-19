import tkinter as tk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from sqlalchemy import create_engine
from selenium.webdriver import ActionChains
import os
import time

# Set the path to the geckodriver executable
geckodriver_path = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\AUTOMATIONS\\geckodriver.exe"

# Set the path to the Firefox binary
firefox_binary = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"

# Set download directory
download_directory = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\DATABASE PROJECT\\Downloads"

# Configure Firefox options
options = Options()
options.binary_location = firefox_binary
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.dir", download_directory)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Initialize the geckodriver service
geckodriver_service = Service(geckodriver_path)

# Initialize a Firefox web driver with the modified options
driver = webdriver.Firefox(service=geckodriver_service, options=options)


# Navigate to the website
driver.get("https://user.repsly.com/account/logon")
wait = WebDriverWait(driver, 20)

language_selector = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '/html/body/repsly-root/div[3]/div[2]/repsly-main-layout/div/div[1]/repsly-login-page/div/div[1]/repsly-language-selector/select')))
select = Select(language_selector)
select.select_by_visible_text("English")


# Find and fill in the email and password fields
email_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="email"]')))
email_field.send_keys("carlos_paz2020@outlook.com")

# Esperar a que el campo de contraseña sea interactuable
password_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]')))
password_field.send_keys("Maxine2021.")

submitbutton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[2]/repsly-main-layout/div/div[1]/repsly-login-page/div/repsly-login-box/div/div[2]/form/div[3]/button')))
submitbutton.click()

time.sleep(5)

driver.get("https://user.repsly.com/reports/form/fd9efcaf-f21a-4c6e-874a-3c4d03d93ce5#summary")

deploymenu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[3]/repsly-main-layout/div/div[2]/repsly-form-report-details-page/div/div[1]/repsly-reports-filters-container/repsly-filters-container/div/div/div[1]/repsly-date-picker-multi-calendar/div/repsly-picker/button/repsly-icon[2]/div/i')))
deploymenu.click()
time.sleep(3)
yesterday = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="cdk-overlay-0"]/div/div/repsly-date-picker-multi-calendar-inline/div/div[1]/repsly-presets/div/ul/li[2]/span')))
yesterday.click()
time.sleep(3)
Applychanges = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app-root"]/div[3]/div[3]/repsly-main-layout/div/div[2]/repsly-form-report-details-page/div/div[1]/repsly-reports-filters-container/repsly-filters-container/div/div/div[3]/repsly-button[2]/button/span/span[1]')))
Applychanges.click()
time.sleep(3)
Export = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app-root"]/div[3]/div[3]/repsly-main-layout/div/div[2]/repsly-form-report-details-page/div/div[1]/repsly-form-report-details-toolbar/div/div[2]/repsly-button/button/span/span[1]')))
Export.click()
time.sleep(3)
Exportexcel = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-dialog-0"]/repsly-dialog-base-content/div/div/repsly-form-report-export-modal/div/repsly-export-selected/repsly-radio-button[1]/label/div[2]/span')))
Exportexcel.click()
time.sleep(3)
Export2 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-dialog-0"]/repsly-dialog-base-content/div/div/repsly-form-report-export-modal/div/div[2]/repsly-button[2]/button/span/span[1]')))
Export2.click()

time.sleep(3)
#CHUNK FOR FTGU WF
driver.get("https://user.repsly.com/reports/form/8df6cdf5-c586-4df8-8a87-99a9d7f9c253#summary")

time.sleep(3)
Export2 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="app-root"]/div[3]/div[3]/repsly-main-layout/div/div[2]/repsly-form-report-details-page/div/div[1]/repsly-form-report-details-toolbar/div/div[2]/repsly-button/button/span/span[1]')))
Export2.click()
time.sleep(3)
Exportexcel2 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-dialog-0"]/repsly-dialog-base-content/div/div/repsly-form-report-export-modal/div/repsly-export-selected/repsly-radio-button[1]/label/div[2]/span')))
Exportexcel2.click()
time.sleep(3)
Export22 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-dialog-0"]/repsly-dialog-base-content/div/div/repsly-form-report-export-modal/div/div[2]/repsly-button[2]/button/span/span[1]')))
Export22.click()
time.sleep(5)

# Close the driver when done
dataframes = []

for filename in os.listdir(download_directory):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(download_directory, filename)
        df = pd.read_excel(file_path, dtype={'Representative ID': str})

        # Eliminar columnas desde la columna 'G' en adelante
        df = df.iloc[:, :6]

        # Cambiar los nombres de las columnas
        df.columns = ["Date", "StoreId", "Place", "Street Address", "Representative ID", "Representative name"]

        # Guardar el archivo modificado
        df.to_excel(file_path, index=False)

        # Agregar el dataframe a la lista
        dataframes.append(df)

# Unir los dataframes en uno solo
all_data = pd.concat(dataframes, ignore_index=True)

# Guardar el archivo combinado
combined_file_path = os.path.join(download_directory, "FTGU All.xlsx")
all_data.to_excel(combined_file_path, index=False)


# Información de conexión a la base de datos
# Configuración de la conexión a la base de datos
server_name = 'Carlos_PC\\SQLEXPRESS'
database_name = 'GothamBrands'
username = 'SA'
password = 'admin123'
connection_string = f'mssql+pyodbc://{username}:{password}@{server_name}/{database_name}?driver=ODBC Driver 17 for SQL Server'
engine = create_engine(connection_string)

# Ruta al archivo Excel
file_path = 'C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\DATABASE PROJECT\\Downloads\\FTGU All.xlsx'

# Lee el archivo Excel, asegurándote de que los IDs sean leídos como strings
df = pd.read_excel(file_path, dtype={'Representative ID': str})

# Corrobora que los datos se hayan leído correctamente
print(df.head())

# Sube los datos a la base de datos, sin modificar los índices
df.to_sql('All_visits', con=engine, if_exists='append', index=False)