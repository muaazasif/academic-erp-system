import os
import sys
from datetime import datetime

# Add the app directory to the path
sys.path.insert(0, '.')

# Import the authentication function
from app import authenticate_google_sheets

def test_google_sheets_connection():
    print("Testing Google Sheets Connection...")
    
    # Check if environment variables are set
    print("\n1. Checking environment variables:")
    google_sheet_id = os.getenv('GOOGLE_SHEET_ID')
    google_sheets_creds = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    
    print(f"   GOOGLE_SHEET_ID: {'SET' if google_sheet_id else 'NOT SET'}")
    print(f"   GOOGLE_SHEETS_CREDENTIALS_JSON: {'SET' if google_sheets_creds else 'NOT SET'}")
    
    if not google_sheet_id or not google_sheets_creds:
        print("\n❌ Environment variables are not properly set in Railway!")
        print("   Please ensure both GOOGLE_SHEET_ID and GOOGLE_SHEETS_CREDENTIALS_JSON are set in Railway dashboard.")
        return False
    
    print("\n2. Attempting to authenticate with Google Sheets API...")
    try:
        service = authenticate_google_sheets()
        if service is None:
            print("❌ Authentication failed - service is None")
            return False
        print("✅ Authentication successful!")
    except Exception as e:
        print(f"❌ Authentication failed with error: {e}")
        return False
    
    print("\n3. Testing write operation to Google Sheet...")
    try:
        # Try to append a test record
        sheet = service.spreadsheets()
        
        # Test data
        test_values = [
            [str(datetime.now().date()), "TEST", "Test User", "12:00:00", "13:00:00", "test", str(datetime.now())]
        ]
        
        body = {'values': test_values}
        
        # Try different ranges
        ranges_to_try = ['Attendance!A:H', 'Sheet1!A:H', 'A:H']
        
        for test_range in ranges_to_try:
            try:
                print(f"   Trying range: {test_range}")
                result = sheet.values().append(
                    spreadsheetId=google_sheet_id,
                    range=test_range,
                    valueInputOption='RAW',
                    body=body
                ).execute()
                
                print(f"   ✅ Successfully wrote to range {test_range}")
                print(f"   Cells updated: {result.get('updates', {}).get('updatedCells', 'Unknown')}")
                return True
                
            except Exception as e:
                print(f"   ❌ Failed to write to {test_range}: {str(e)}")
                continue
        
        print("❌ Failed to write to any range")
        return False
        
    except Exception as e:
        print(f"❌ Error during write test: {e}")
        return False

if __name__ == "__main__":
    success = test_google_sheets_connection()
    if success:
        print("\n🎉 Google Sheets integration is working correctly!")
    else:
        print("\n💥 Google Sheets integration needs to be fixed.")