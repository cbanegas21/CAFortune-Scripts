import tkinter as tk
from msal import PublicClientApplication
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import win32com
import win32com.client as win32
import pyexcel as p
from sqlalchemy import create_engine
from selenium.webdriver import ActionChains
import os
import time
from selenium.webdriver.common.action_chains import ActionChains


# Set the path to the geckodriver executable
geckodriver_path = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\AUTOMATIONS\\geckodriver.exe"

# Set the path to the Firefox binary
firefox_binary = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"

# Set download directory
download_directory = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\DATA FOR VISITS"

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

time.sleep(3)
#CHUNK FOR extraction
driver.get("https://user.repsly.com/manage/export")

FORMSCLICK = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportFormsBtn"]')))
FORMSCLICK.click()

elements = [
    "FTGU-WF",
    "FTGU-WF",
    "FTGU-All Retail - 1/3",
]
fecha_inicio = "04/09/2024"  # Ajusta estas fechas según necesites
fecha_fin = "04/14/2024"

campo_fecha_inicio = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateBegin"]')))
campo_fecha_inicio.clear()
campo_fecha_inicio.send_keys(fecha_inicio)

campo_fecha_fin = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateEnd"]')))
campo_fecha_fin.clear()
campo_fecha_fin.send_keys(fecha_fin)




#safe_click_target = WebDriverWait(driver, 10).until(
#    EC.element_to_be_clickable((By.XPATH, '//body'))
#)
#driver.execute_script("arguments[0].click();", safe_click_target)


for elemento in elements:
    # Encuentra el dropdown y crea el objeto Select
    dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "rn-export-forms-dropdown")))
    select = Select(dropdown)
    # Selecciona el elemento deseado usando el texto visible
    select.select_by_visible_text(elemento)

    campofiletype = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "rn-forms-exportType")))
    select = Select(campofiletype)
    select.select_by_visible_text("Excel")

    # Localiza el botón por su ID
    xpath_del_boton = "/html/body/repsly-root/div[3]/div[2]/main-layout/div/div[6]/div/div/div/ng-transclude/div/div[3]/div/div[2]/ul/li[8]/div/table/tbody/tr[4]/td[2]/a/span"

    try:
        # Espera a que el botón sea detectado como clickeable en el DOM
        boton_descargar = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, xpath_del_boton))
        )
        # Hacer clic en el botón directamente
        boton_descargar.click()
        time.sleep(5)
    except Exception as e:
        print(f"Error al intentar hacer clic en el botón: {e}")
    # Espera entre descargas para evitar sobrecargar el servidor o ser detectado como un bot
    time.sleep(5)


download_to_brand_mapping = {
    "Forms_FTGU-All_Retail_1_3": "FTGU All Retail",
    "Forms_FTGU-WF": "FTGU Whole Foods"
}

# Function to rename the downloaded files
for old_name, new_name in download_to_brand_mapping.items():
    old_file_path = os.path.join(download_directory, f"{old_name}.xlsx")
    new_file_path = os.path.join(download_directory, f"{new_name}.xlsx")

    # Check if the old file exists before renaming
    if os.path.exists(old_file_path):
        os.rename(old_file_path, new_file_path)
        print(f"Renamed {old_name} to {new_name}.xlsx")
    else:
        print(f"File {old_name}.xlsx does not exist. Skipping.")

def convert_excel_files_in_folder(folder_path):
    try:
        excel = win32com.client.DispatchEx('Excel.Application')
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx"):  # Adjust this condition for other Excel formats if needed
                full_path = os.path.join(folder_path, filename)
                print(f"Opening and saving: {filename}")
                workbook = excel.Workbooks.Open(full_path)
                # Save the workbook, overwriting the original file
                workbook.Save()
                workbook.Close()
        excel.Quit()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        print("All files have been processed.")


# Assuming 'download_directory' is where your files are downloaded and renamed
download_directory = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\DATA FOR VISITS"
convert_excel_files_in_folder(download_directory)

# Assuming 'download_directory' is where your files are downloaded and renamed
file_path1 = os.path.join(download_directory, "FTGU All Retail.xlsx")
file_path2 = os.path.join(download_directory, "FTGU Whole Foods.xlsx")

# Function to clean the data as specified
def clean_data(df):
    # Delete everything from column "I" and beyond
    df = df.iloc[:, :8]  # Keep only the first 8 columns (A-H)
    # Delete columns "F" and "G"
    df.drop(columns=[df.columns[5], df.columns[6]], inplace=True)
    # Change the header of column "D" to "StoreId"
    df.rename(columns={df.columns[3]: 'StoreId'}, inplace=True)
    return df

# Read the Excel files
df1 = pd.read_excel(file_path1)
df2 = pd.read_excel(file_path2)

# Clean both DataFrames
df1_cleaned = clean_data(df1)
df2_cleaned = clean_data(df2)

# Merge both cleaned DataFrames
merged_df = pd.concat([df1_cleaned, df2_cleaned], ignore_index=True)

# Optional: Save the merged DataFrame to a new Excel file
merged_file_path = os.path.join(download_directory, "RepVisits.xlsx")
merged_df.to_excel(merged_file_path, index=False)

print("Data cleaning and merging complete.")


#------------------------------------------------------*
# Loop through all files in the download directory
for filename in os.listdir(download_directory):
    file_path = os.path.join(download_directory, filename)
    
    # Check if it's a file and not the 'RepVisits.xlsx'
    if os.path.isfile(file_path) and filename != "RepVisits.xlsx":
        os.remove(file_path)  # Delete the file
        print(f"Deleted {filename}")

print("Cleanup complete.")