import urllib.request
import json
import os

url = "https://muaazasif-4cf8e-default-rtdb.firebaseio.com/.json"

print(f"Fetching Firebase data from {url}...")
try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as response:
        if response.status == 200:
            data = json.loads(response.read().decode('utf-8'))
            if data is None:
                print("Firebase returned null/empty data.")
            else:
                print("Successfully fetched Firebase data!")
                print("Root keys in Firebase:", list(data.keys()))
                for key, val in data.items():
                    if isinstance(val, dict):
                        print(f"  - Key: {key}, entries: {len(val)}")
                    else:
                        print(f"  - Key: {key}, type: {type(val)}")
                
                # Save to file
                with open("firebase_dump.json", "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                print("Saved dump to firebase_dump.json")
        else:
            print(f"Failed to fetch. Status: {response.status}")
except Exception as e:
    print(f"Error fetching data: {e}")
