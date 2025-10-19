import requests
import json
import datetime
import re
from requests.auth import HTTPBasicAuth

# Credentials
username = '80941603-F785-4E0F-8AB1-ED798E54F88C'
password = '17BCD0FD-94DD-4B0C-B059-76D68C1145A8'

# API setup
base_url = 'https://api.repsly.com/v3/export/photos/'
headers = {
    'Accept': 'application/json',
}

# Helper to parse Repsly's date format
def parse_repsly_date(date_str):
    match = re.search(r'/Date\((\d+)', date_str)
    if match:
        timestamp_ms = int(match.group(1))
        return datetime.datetime.fromtimestamp(timestamp_ms / 1000).strftime('%Y-%m-%d %H:%M:%S')
    return "Invalid Date"

# Start from the beginning
last_photo_id = 0
all_photos = []

print("Starting photo export...\n")

while True:
    url = f'{base_url}{last_photo_id}'
    response = requests.get(url, headers=headers, auth=HTTPBasicAuth(username, password))
    
    if response.status_code != 200:
        print(f"Error {response.status_code}: {response.text}")
        break
    
    data = response.json()
    photos = data.get('Photos', [])
    total_count = data.get('MetaCollectionResult', {}).get('TotalCount', 0)
    first_id = data.get('MetaCollectionResult', {}).get('FirstID')
    last_id = data.get('MetaCollectionResult', {}).get('LastID')

    if photos:
        # Log batch date
        first_photo_raw_date = photos[0].get('DateAndTime', '')
        first_photo_date = parse_repsly_date(first_photo_raw_date)
        print(f"Fetched {len(photos)} photos starting from ID {first_id} (Date: {first_photo_date})")

        # Add parsed date to each photo
        for photo in photos:
            raw_date = photo.get('DateAndTime', '')
            photo['DateAndTimeParsed'] = parse_repsly_date(raw_date)

        # Append batch
        all_photos.extend(photos)

        # Prepare for next batch
        last_photo_id = last_id
    else:
        print("No more photos found.")
        break

    if total_count == 0:
        break

# Save to JSON
tagged_photos = [photo for photo in all_photos if photo.get('Tag')]

# Save filtered data to JSON
output_file = 'Tagged_Photos.json'
with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(tagged_photos, f, ensure_ascii=False, indent=2)

print(f"\nDone! Saved {len(tagged_photos)} tagged photos to {output_file}")
