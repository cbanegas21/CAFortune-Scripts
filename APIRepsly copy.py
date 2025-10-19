import requests
from requests.auth import HTTPBasicAuth
import time
import json
from datetime import datetime

# User credentials
username = '80941603-F785-4E0F-8AB1-ED798E54F88C'
password = '17BCD0FD-94DD-4B0C-B059-76D68C1145A8'

# API Configuration
base_url = 'https://api.repsly.com/v3/export/forms/'
headers = {'Content-Type': 'application/json'}

# Start with the last known LastID (modify if needed)
last_form_id = 238574406  
all_forms = []  # Store all retrieved records

# Maximum retries
max_retries = 5
retry_delay = 5  # Seconds between retries

# Helper function to convert the date format
def convert_ms_date(ms_date_string):
    timestamp = int(ms_date_string[6:19])  
    return datetime.utcfromtimestamp(timestamp / 1000).strftime('%Y-%m-%d %H:%M:%S')

# Fetch all records
while True:
    total_count = 0  # Track total records
    
    for attempt in range(max_retries):
        try:
            url = f'{base_url}{last_form_id}'
            response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))
            
            if response.status_code == 200:
                data = response.json()
                forms = data.get('Forms', [])
                all_forms.extend(forms)  # Store retrieved records

                if forms and 'DateAndTime' in forms[0]:
                    readable_date = convert_ms_date(forms[0]['DateAndTime'])
                    print(f"Form Date: {readable_date}")

                total_count = data['MetaCollectionResult']['TotalCount']
                last_form_id = data['MetaCollectionResult']['LastID']  

                print(f"{len(forms)} forms retrieved. LastID updated to {last_form_id}.")
                
                break  # Exit retry loop if successful
            else:
                print(f"Error {response.status_code}: {response.text}")
                raise Exception(f"HTTP error {response.status_code}")

        except Exception as e:
            print(f"Attempt {attempt+1}/{max_retries} failed: {e}")
            if attempt + 1 == max_retries:
                print(f"Max retries reached. Stopping at LastID: {last_form_id}")
                break
            time.sleep(retry_delay)

    if total_count == 0 or attempt + 1 == max_retries:
        break  # Exit main loop when no more data

# Save all data to a single JSON file
output_file = "all_forms.json"
with open(output_file, "w", encoding="utf-8") as f:
    json.dump(all_forms, f, indent=4)

print(f"\nSuccessfully saved {len(all_forms)} records to {output_file}!")  # Remove emoji

