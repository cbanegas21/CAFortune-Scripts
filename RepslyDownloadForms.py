import tkinter as tk
from tkinter import ttk, messagebox
from selenium import webdriver
from tkcalendar import DateEntry
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk
import threading
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import win32com.client
import shutil
from pywinauto.application import Application
from openpyxl import load_workbook

# Set the path to the geckodriver executable
geckodriver_path = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\AUTOMATIONS\\geckodriver.exe"

# Set the path to the Firefox binary
firefox_binary = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"

# Set download directory
download_directory = "C:\\Users\\carlo\\OneDrive\\Escritorio\\Gotham 2023-2024\\DATA FOR REPORTING"

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
    
    language_dropdown = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@name="languages"]'))
    )
    select = Select(language_dropdown)
    select.select_by_visible_text("English")
    
    
    email_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="email"]')))
    email_field.send_keys("carlos_paz2020@outlook.com")
    nextbutton = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[2]/repsly-main-layout/div/div[1]/repsly-login-page/div/repsly-login-box/div/div[2]/form/div[2]/button')))
    nextbutton.click()
    password_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="okta-signin-password"]')))
    password_field.send_keys("Maxine2023.")
    time.sleep(3)
    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="okta-signin-submit"]')))
    submit_button.click()


# Function to navigate to the export page
def navigate_to_export_page(driver):
    
    search1 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/repsly-root/div[3]/div[3]/repsly-main-layout/div/div[1]/repsly-top-line-widget/div/div/div/div[2]/div[3]/repsly-settings-widget/div/div/repsly-button/button/span/span/repsly-icon/div/i')))
    search1.click()
    search2 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[9]/div/div/div/div/repsly-settings-menu/div/div/repsly-settings-menu-section[2]/div/ul/li[2]')))
    search2.click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportFormsBtn"]'))).click()

# Function to extract dropdown options using Selenium
def get_dropdown_options(driver):
    # Extract the dropdown options
    dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "rn-export-forms-dropdown")))
    select = Select(dropdown)
    options = [option.text for option in select.options]

    return options

# Function to perform export based on selected forms
def run_export(driver, selected_forms, fecha_inicio, fecha_fin):
    print("Starting export process...")

    for form_name in selected_forms:
        print(f"Processing form: {form_name}")
        dropdown = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "rn-export-forms-dropdown")))
        select = Select(dropdown)
        select.select_by_visible_text(form_name)

        file_type_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "rn-forms-exportType")))
        select = Select(file_type_field)
        select.select_by_visible_text("Excel")

        start_date_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateBegin"]')))
        start_date_field.clear()
        start_date_field.send_keys(fecha_inicio)

        end_date_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="FormsDateEnd"]')))
        end_date_field.clear()
        end_date_field.send_keys(fecha_fin)

        download_button_xpath = "/html/body/repsly-root/div[3]/div[2]/main-layout/div/div[6]/div/div/div/ng-transclude/div/div[3]/div/div[2]/ul/li[8]/div/table/tbody/tr[4]/td[2]/a/span"

        try:
            download_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, download_button_xpath)))
            download_button.click()
            print(f"Clicked download button for form: {form_name}")
            time.sleep(15)  # Wait for the download to complete
        except Exception as e:
            print(f"Error clicking the download button: {e}")

        time.sleep(2)  # Wait between downloads to avoid overloading the server or being detected as a bot

        list_of_files = os.listdir(download_directory)
        full_path = [os.path.join(download_directory, f) for f in list_of_files]
        latest_file = max(full_path, key=os.path.getctime)

        new_filename = os.path.join(download_directory, f"{form_name}.xlsx")
        shutil.move(latest_file, new_filename)
        print(f"Renamed file to: {new_filename}")

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
                        print(f"Deleted column: {column}")
                except Exception as e:
                    print(f"Error deleting column {column}: {e}")

            workbook.Save()
            workbook.Close(False)
            excel_app.Quit()
            
            print(f"Opened, modified, saved, and closed the file in Excel: {new_filename}")
        except Exception as e:
            print(f"Error opening, modifying, saving, or closing the file in Excel: {e}")

        time.sleep(3)  # Wait to ensure the file is closed

    print("Export process completed successfully.")



def create_gui(driver):
    root = tk.Tk()
    root.title("Select Forms to Export")

    selected_options = {}
    all_options = get_dropdown_options(driver)

    def show_loading():
        pass  # Placeholder function, no loading indicator

    def hide_loading():
        pass  # Placeholder function, no loading indicator

    def on_export():
        selected_forms = [option for option, var in selected_options.items() if var.get()]
        if not selected_forms:
            messagebox.showerror("Error", "No forms selected!")
            return

        fecha_inicio = start_date_entry.get()
        fecha_fin = end_date_entry.get()

        show_loading()
        root.update_idletasks()  # Update the GUI to show the loading indicator
        run_export(driver, selected_forms, fecha_inicio, fecha_fin)
        hide_loading()
        messagebox.showinfo("Success", "Export completed!")

    def on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        selected_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def filter_options(event):
        search_term = search_var.get().lower()
        for widget in scrollable_frame.winfo_children():
            widget.destroy()
        for option in all_options:
            if search_term in option.lower():
                if option not in selected_options:
                    selected_options[option] = tk.BooleanVar()
                chk = ttk.Checkbutton(scrollable_frame, text=option, variable=selected_options[option], command=update_selected_list)
                chk.pack(anchor="w")

    def update_selected_list():
        selected_items = [option for option, var in selected_options.items() if var.get()]
        selected_count.set(f"Selected: {len(selected_items)}")
        for widget in selected_scrollable_frame.winfo_children():
            widget.destroy()
        for item in selected_items:
            tk.Label(selected_scrollable_frame, text=item).pack(anchor="w")

    frame = ttk.Frame(root, padding="10")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    # Date inputs using DateEntry
    tk.Label(frame, text="Start Date:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    start_date_entry = DateEntry(frame, date_pattern='mm/dd/yyyy')
    start_date_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

    tk.Label(frame, text="End Date:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    end_date_entry = DateEntry(frame, date_pattern='mm/dd/yyyy')
    end_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

    search_var = tk.StringVar()
    search_box = ttk.Entry(frame, textvariable=search_var)
    search_box.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
    search_box.bind('<KeyRelease>', filter_options)

    canvas = tk.Canvas(frame)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    for option in all_options:
        selected_options[option] = tk.BooleanVar()
        chk = ttk.Checkbutton(scrollable_frame, text=option, variable=selected_options[option], command=update_selected_list)
        chk.pack(anchor="w")

    canvas.grid(row=3, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
    scrollbar.grid(row=3, column=2, sticky=(tk.N, tk.S))

    btn_export = ttk.Button(frame, text="Export", command=on_export)
    btn_export.grid(row=4, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

    # Selected items and count
    selected_count = tk.StringVar(value="Selected: 0")
    tk.Label(frame, textvariable=selected_count).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)

    selected_canvas = tk.Canvas(frame)
    selected_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=selected_canvas.yview)
    selected_scrollable_frame = ttk.Frame(selected_canvas)

    selected_scrollable_frame.bind(
        "<Configure>",
        lambda e: selected_canvas.configure(
            scrollregion=selected_canvas.bbox("all")
        )
    )

    selected_canvas.create_window((0, 0), window=selected_scrollable_frame, anchor="nw")
    selected_canvas.configure(yscrollcommand=selected_scrollbar.set)

    selected_canvas.grid(row=1, column=2, rowspan=4, padx=5, pady=5, sticky=(tk.N, tk.W, tk.E, tk.S))
    selected_scrollbar.grid(row=1, column=3, rowspan=4, sticky=(tk.N, tk.S))

    # Enable mouse scrolling
    canvas.bind_all("<MouseWheel>", on_mousewheel)
    selected_canvas.bind_all("<MouseWheel>", on_mousewheel)

    root.mainloop()


# Main execution
driver = initialize_webdriver()
try:
    login_to_website(driver)
    navigate_to_export_page(driver)
    create_gui(driver)
finally:
    driver.quit()
