import os
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get the credentials JSON from environment
creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')

if creds_json:
    # Parse and fix the private key
    credentials_info = json.loads(creds_json)
    
    # Fix the private key by replacing literal \n with actual newlines
    if 'private_key' in credentials_info:
        private_key = credentials_info['private_key']
        # Replace escaped newlines with actual newlines
        private_key = private_key.replace('\\n', '\n')
        credentials_info['private_key'] = private_key

    # Write to credentials.json file
    with open('credentials.json', 'w') as f:
        json.dump(credentials_info, f, indent=2)
    
    print("Created credentials.json file with properly formatted private key")
else:
    print("GOOGLE_SHEETS_CREDENTIALS_JSON not found in environment")