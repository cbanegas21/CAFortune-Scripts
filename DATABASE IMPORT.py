import pandas as pd
import json
import os

def process_json_to_excel(json_path, output_folder):
    """
    Reads a JSON file, extracts structured data based on FormName,
    and saves each form as a separate Excel file.
    """
    # Load JSON data
    with open(json_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
    
    # Create output folder if not exists
    os.makedirs(output_folder, exist_ok=True)
    
    form_tables = {}
    
    for entry in data:
        form_name = entry["FormName"].replace(" ", "_").replace("/", "_").replace("-", "_")  # Clean table name
        
        if form_name not in form_tables:
            form_tables[form_name] = []
        
        # Extract top-level fields
        row_data = {
            "FormID": entry.get("FormID"),
            "ClientCode": entry.get("ClientCode"),
            "ClientName": entry.get("ClientName"),
            "DateAndTime": entry.get("DateAndTime"),
            "RepresentativeCode": entry.get("RepresentativeCode"),
            "RepresentativeName": entry.get("RepresentativeName"),
            "StreetAddress": entry.get("StreetAddress"),
            "ZIP": entry.get("ZIP"),
            "City": entry.get("City"),
            "State": entry.get("State"),
            "Country": entry.get("Country"),
            "Email": entry.get("Email"),
            "Phone": entry.get("Phone"),
            "Mobile": entry.get("Mobile"),
            "Territory": entry.get("Territory"),
            "Longitude": entry.get("Longitude"),
            "Latitude": entry.get("Latitude"),
            "SignatureURL": entry.get("SignatureURL"),
            "VisitStart": entry.get("VisitStart"),
            "VisitEnd": entry.get("VisitEnd"),
            "VisitID": entry.get("VisitID"),
        }
        
        # Extract Items as additional columns, excluding overly long fields
        for item in entry.get("Items", []):
            field = item["Field"].strip().replace(" ", "_").replace("/", "_").replace("-", "_")  # Clean column name
            value = item["Value"]
            if len(field) <= 255:  # Exclude fields exceeding character limit
                row_data[field] = value

        form_tables[form_name].append(row_data)
    
    # Save each form data as an Excel file
    excel_files = []
    for form, rows in form_tables.items():
        df = pd.DataFrame(rows)
        excel_path = os.path.join(output_folder, f"{form}.xlsx")
        df.to_excel(excel_path, index=False)
        excel_files.append(excel_path)
    
    return excel_files

# Updated paths
json_path = r"C:\Users\carlo\OneDrive\Escritorio\Gotham 2023-2024\DATABASE PROJECT\all_forms.json"  
output_folder = r"C:\Users\carlo\OneDrive\Escritorio\Gotham 2023-2024\DATABASE PROJECT\DATABASE_CAFORTUNE"

excel_files = process_json_to_excel(json_path, output_folder)
print(f"Excel files saved: {excel_files}")
