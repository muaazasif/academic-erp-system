import os
import json
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def test_private_key_format():
    """Test the private key format"""
    creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    if not creds_json:
        print("ERROR: GOOGLE_SHEETS_CREDENTIALS_JSON not found")
        return
    
    try:
        credentials_info = json.loads(creds_json)
        private_key = credentials_info.get('private_key', '')
        
        print("Original private key (first 100 chars):")
        print(repr(private_key[:100]))
        print()
        
        # Apply the same transformation as in the app
        fixed_private_key = private_key.replace('\\n', '\n')
        
        print("Fixed private key (first 100 chars):")
        print(repr(fixed_private_key[:100]))
        print()
        
        print("Does it contain proper newlines now?", '\\n' not in fixed_private_key and '\n' in fixed_private_key)
        
        # Show a few lines of the key to verify formatting
        lines = fixed_private_key.split('\n')
        print(f"\nNumber of lines in private key: {len(lines)}")
        print("First few lines:")
        for i, line in enumerate(lines[:10]):  # Show first 10 lines
            print(f"  {i}: {repr(line)}")
        
    except Exception as e:
        print(f"Error parsing credentials: {e}")

if __name__ == "__main__":
    test_private_key_format()