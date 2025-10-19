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

# LastID que se quiere obtener
last_form_id = 236800769  # Reemplaza con el LastID que deseas obtener

# Helper function to convert the date from /Date(1698045694000+0000)/ format
def convert_ms_date(ms_date_string):
    timestamp = int(ms_date_string[6:19])  # Extract the Unix timestamp (in milliseconds)
    return datetime.utcfromtimestamp(timestamp / 1000).strftime('%Y-%m-%d %H:%M:%S')

# Make the GET request with authentication
url = f'{base_url}{last_form_id}'
response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))

# If the response is successful
if response.status_code == 200:
    data = response.json()
    
    # Save the JSON to a file
    with open('last_form.json', 'w') as json_file:
        json.dump(data, json_file, indent=4)
    
    print(f"El JSON del LastID {last_form_id} ha sido guardado en 'last_form.json'.")
else:
    print(f"Error {response.status_code}: {response.text}")