import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import os
import time
import logging
import shutil
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Configuration variables
geckodriver_path = "C:\\Users\\carlo\\OneDrive\\Documents\\ALL PROJECTS\\geckodriver.exe"
firefox_binary = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"
download_directory = "C:\\Users\\carlo\\OneDrive\Documents\\ALL PROJECTS\\CA Fortune\\DATA IMPORTS"
form_names_file = "C:\\Users\\carlo\\OneDrive - C.A. Fortune - C.A. Carlin\\Gotham Dashboards\\Regional & Store Level Trackers\\form_names.txt"

# Configure Firefox options
def create_firefox_options():
    options = Options()
    options.binary_location = firefox_binary
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", download_directory)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return options

# Initialize the geckodriver service
geckodriver_service = Service(geckodriver_path)

# Function to initialize the WebDriver
def initialize_webdriver():
    options = create_firefox_options()
    driver = webdriver.Firefox(service=geckodriver_service, options=options)
    return driver

# Function to login to the website
def login_to_website(driver):
    driver.get("https://user.repsly.com/manage/export")
    
    # Select language
    language_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@name="languages"]'))
    )
    select = Select(language_dropdown)
    select.select_by_visible_text("English")
    
    # Enter email
    email_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="email"]')))
    email_field.send_keys("carlos_paz2020@outlook.com")
    
    # Click next
    next_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[2]/repsly-main-layout/div/div[1]/repsly-login-page/div/repsly-login-box/div/div[2]/form/div[2]/button')))
    next_button.click()
    
    # Enter password
    password_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="okta-signin-password"]')))
    password_field.send_keys("Maxine2023.")
    
    # Submit
    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="okta-signin-submit"]')))
    submit_button.click()

# Function to navigate to the export page
def navigate_to_export_page(driver):
    # Click settings icon
    settings_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[3]/repsly-main-layout/div/div[1]/repsly-top-line-widget/div/div/div/div[2]/div[3]/repsly-settings-widget/div/div/repsly-button/button/span/span/repsly-icon/div/i')))
    settings_icon.click()
    
    # Select export option
    export_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div/div/div/div/repsly-settings-menu/div/div/repsly-settings-menu-section[2]/div/ul/li[2]')))
    export_option.click()
    
    # Click export forms button
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportFormsBtn"]'))).click()

# Function to read form names from a text file
def read_form_names(file_path):
    with open(file_path, 'r') as file:
        form_names = file.read().splitlines()
    return form_names

# Function to perform export based on selected forms
def run_export(driver, selected_forms, fecha_inicio, fecha_fin):
    logging.info("Starting export process...")
    failed_forms = []

    for form_name in selected_forms:
        logging.info(f"Processing form: {form_name}")
        
        try:
            # Select form
            dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "rn-export-forms-dropdown")))
            select = Select(dropdown)
            select.select_by_visible_text(form_name)

            # Select file type
            file_type_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "rn-forms-exportType")))
            select = Select(file_type_field)
            select.select_by_visible_text("Excel")

            # Set date range
            start_date_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateBegin"]')))
            start_date_field.clear()
            start_date_field.send_keys(fecha_inicio)

            end_date_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateEnd"]')))
            end_date_field.clear()
            end_date_field.send_keys(fecha_fin)

            # Click download button
            download_button_xpath = "/html/body/repsly-root/div[3]/div[2]/main-layout/div/div[6]/div/div/div/ng-transclude/div/div[3]/div/div[2]/ul/li[8]/div/table/tbody/tr[4]/td[2]/a/span"

            download_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, download_button_xpath)))
            download_button.click()
            logging.info(f"Clicked download button for form: {form_name}")
            time.sleep(15)  # Wait for the download to complete
        except Exception as e:
            logging.error(f"Error processing form {form_name}: {e}")
            failed_forms.append(form_name)
            continue

        time.sleep(2)  # Wait between downloads to avoid overloading the server or being detected as a bot

        # Move and rename the downloaded file
        list_of_files = os.listdir(download_directory)
        full_path = [os.path.join(download_directory, f) for f in list_of_files]
        latest_file = max(full_path, key=os.path.getctime)

        new_filename = os.path.join(download_directory, f"{form_name}.xlsx")
        shutil.move(latest_file, new_filename)
        logging.info(f"Renamed file to: {new_filename}")

        # Modify the Excel file
        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            workbook = excel_app.Workbooks.Open(new_filename)
            worksheet = workbook.Worksheets(1)

            # Delete specified columns
            columns_to_delete = ["Custom fields.On Premise", "Custom fields.Primary Distributor", "Custom fields.Secondary Distributor", "Custom fields.Custom Label"]
            for column in columns_to_delete:
                try:
                    column_index = None
                    for col in range(1, worksheet.UsedRange.Columns.Count + 1):
                        if worksheet.Cells(1, col).Value == column:
                            column_index = col
                            break
                    if column_index:
                        worksheet.Columns(column_index).Delete()
                        logging.info(f"Deleted column: {column}")
                except Exception as e:
                    logging.error(f"Error deleting column {column}: {e}")

            workbook.Save()
            workbook.Close(False)
            excel_app.Quit()
            
            logging.info(f"Opened, modified, saved, and closed the file in Excel: {new_filename}")
        except Exception as e:
            logging.error(f"Error opening, modifying, saving, or closing the file in Excel: {e}")
            failed_forms.append(form_name)

        time.sleep(3)  # Wait to ensure the file is closed

    logging.info("Export process completed.")
    if failed_forms:
        logging.warning(f"The following forms were not downloaded or processed correctly: {', '.join(failed_forms)}")

# Function to create and display the date entry UI
def create_date_entry_ui():
    def on_submit():
        fecha_inicio = start_date_entry.get()
        fecha_fin = end_date_entry.get()
        root.destroy()  # Close the UI

        # Start the export process with the entered dates
        main(fecha_inicio, fecha_fin)

    root = tk.Tk()
    root.title("Enter Date Range")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    tk.Label(frame, text="Start Date:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    start_date_entry = DateEntry(frame, date_pattern='mm/dd/yyyy')
    start_date_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

    tk.Label(frame, text="End Date:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    end_date_entry = DateEntry(frame, date_pattern='mm/dd/yyyy')
    end_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

    submit_button = ttk.Button(frame, text="Submit", command=on_submit)
    submit_button.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

    root.mainloop()

# Main execution
def main(fecha_inicio, fecha_fin):
    # Read form names from the text file
    form_names = read_form_names(form_names_file)

    driver = initialize_webdriver()
    try:
        login_to_website(driver)
        navigate_to_export_page(driver)
        run_export(driver, form_names, fecha_inicio, fecha_fin)
    finally:
        driver.quit()

if __name__ == "__main__":
    create_date_entry_ui()
