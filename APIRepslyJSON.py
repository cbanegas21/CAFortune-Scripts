import requests
from requests.auth import HTTPBasicAuth
import time
from datetime import datetime
import json

# User credentials
username = '80941603-F785-4E0F-8AB1-ED798E54F88C'
password = '17BCD0FD-94DD-4B0C-B059-76D68C1145A8'

# Initial configuration
base_url = 'https://api.repsly.com/v3/export/forms/'
headers = {
    'Content-Type': 'application/json'
}

# Start with the LastID where the error occurred
last_form_id = 246101970  # Example LastID (replace with your last one)
all_forms = []

# Maximum number of retries
max_retries = 5
retry_delay = 5  # Seconds to wait between retries

# Helper function to convert the date from /Date(1698045694000+0000)/ format
def convert_ms_date(ms_date_string):
    timestamp = int(ms_date_string[6:19])  # Extract the Unix timestamp (in milliseconds)
    return datetime.utcfromtimestamp(timestamp / 1000).strftime('%Y-%m-%d %H:%M:%S')

# Make requests until there are no more forms
while True:
    total_count = 0  # Initialize total_count before the loop
    
    for attempt in range(max_retries):
        try:
            # Build the URL for the request with the current LastID
            url = f'{base_url}{last_form_id}'
            
            # Make the GET request with authentication
            response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))
            
            # If the response is successful
            if response.status_code == 200:
                data = response.json()
                
                # Check if there are any Forms entries
                forms = data.get('Forms', [])
                all_forms.extend(forms)
                
                # Print only metadata for the current batch
                total_count = data['MetaCollectionResult']['TotalCount']
                last_form_id = data['MetaCollectionResult']['LastID']  # Update the LastID
                
                print(f"{len(forms)} forms were retrieved in this request. LastID updated to {last_form_id}.")
                
                # If there are no more forms (TotalCount is 0), exit the loop
                if total_count == 0:
                    print("All available forms have been retrieved.")
                    break
                
                # Reset the retry counter if the request was successful
                break
            else:
                print(f"Error {response.status_code}: {response.text}")
                raise Exception(f"HTTP error {response.status_code}")
        
        except Exception as e:
            print(f"Attempt {attempt+1}/{max_retries} failed: {e}")
            
            # If it's the last attempt, exit the loop
            if attempt + 1 == max_retries:
                print(f"Maximum number of retries reached. Last LastID: {last_form_id}")
                break
            
            # Wait before retrying
            time.sleep(retry_delay)
    
    # If total_count is 0 or the maximum number of retries is reached, stop the loop
    if total_count == 0 or attempt + 1 == max_retries:
        break

# Save all forms to a JSON file
with open('retrieved_forms.json', 'w') as json_file:
    json.dump(all_forms, json_file, indent=4)

print(f"A total of {len(all_forms)} forms were retrieved and saved to 'retrieved_forms.json'.")
