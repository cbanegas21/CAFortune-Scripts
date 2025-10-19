import requests
import json


encoded_credentials = 'ODA5NDE2MDMtRjc4NS00RTBGLThBQjEtRUQ3OThFNTRGODhDOjE3QkNEMEZELTk0REQtNEIwQy1CMDU5LTc2RDY4QzExNDVBOA=='


url = "https://api.repsly.com/v3/export/dailyworkingtime/3338892"


headers = {
    "Authorization": f"Basic {encoded_credentials}",
    "Content-Type": "application/json"
}


response = requests.get(url, headers=headers)


if response.status_code == 200:
    
    with open('api_response.json', 'w') as json_file:
        json.dump(response.json(), json_file, indent=4)
    print("Success! Data retrieved and saved to 'api_response2.json'.")
else:
    print(f"Failed to retrieve data. Status code: {response.status_code}")
