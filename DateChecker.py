import os
import pandas as pd
from tkinter import Tk, filedialog
import warnings

# Suppress the user warning for date parsing
warnings.filterwarnings("ignore", message="Could not infer format")

def get_most_recent_date(file_path):
    try:
        df = pd.read_excel(file_path)
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            most_recent_date = df['Date'].max()
            return most_recent_date.date()  # Remove the time part and return only the date
        else:
            return None
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None

def main():
    # Initialize Tkinter and hide the root window
    root = Tk()
    root.withdraw()

    # Ask the user to select a folder
    folder_selected = filedialog.askdirectory(title="Select Folder with Excel Files")

    if not folder_selected:
        print("No folder selected, exiting...")
        return

    results = []

    # Iterate over all Excel files in the folder
    for filename in os.listdir(folder_selected):
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            file_path = os.path.join(folder_selected, filename)
            most_recent_date = get_most_recent_date(file_path)
            if most_recent_date:
                results.append((filename, most_recent_date))

    # Sort results by date (oldest to newest)
    results.sort(key=lambda x: x[1])

    # Print the sorted results
    for filename, date in results:
        print(f"{filename}: Most Recent Date - {date}")

if __name__ == "__main__":
    main()
