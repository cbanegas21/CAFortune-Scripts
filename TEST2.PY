import pandas as pd

# The path to your JSON file
json_file_path = 'api_response.json'

# The path where you want to save the Excel file
excel_file_path = 'converted_data.xlsx'

# Load the JSON data into a pandas DataFrame
df = pd.read_json(json_file_path)

# Convert the DataFrame to an Excel file
df.to_excel(excel_file_path, index=False, engine='openpyxl')

print(f'JSON data has been converted to Excel and saved to {excel_file_path}')
