import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import warnings

# Suppress specific date parsing warnings
warnings.filterwarnings("ignore", message="Could not infer format, so each element will be parsed individually, falling back to `dateutil`")

# Hide the root Tkinter window
root = tk.Tk()
root.withdraw()

# Ask the user to select a folder
folder_path = filedialog.askdirectory(title="Select Folder with Excel Files")
if not folder_path:
    print("No folder selected. Exiting...")
    exit()

# List to store the names of Excel files containing 2024 data
excel_files_with_2024 = []

# Iterate through each file in the selected folder
for file in os.listdir(folder_path):
    if file.endswith(('.xlsx', '.xls')):
        file_path = os.path.join(folder_path, file)
        try:
            # Read the Excel file (default: first sheet)
            df = pd.read_excel(file_path)
            
            # Check if the 'Date' column exists
            if 'Date' in df.columns:
                # Convert the 'Date' column to datetime.
                # Optionally specify the format if known, e.g., format='%m/%d/%Y'
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
                
                # Filter rows where the year is 2024
                df_2024 = df[df['Date'].dt.year == 2024]
                
                # If there is at least one row with a 2024 date, add the file to the list
                if not df_2024.empty:
                    excel_files_with_2024.append(file)
            else:
                print(f"'Date' column not found in {file}. Skipping this file.")
        except Exception as e:
            print(f"Error reading {file}: {e}")

# Output the list of Excel files that contain 2024 data
print("Excel files containing 2024 data:")
for file in excel_files_with_2024:
    print(file)
