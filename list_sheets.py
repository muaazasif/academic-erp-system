import os
import json
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build

# Load environment variables
load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def list_spreadsheet_sheets():
    """List all sheets in the spreadsheet"""
    try:
        print("Listing sheets in the spreadsheet...")
        
        # Get credentials from environment variable
        creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
        if not creds_json:
            print("ERROR: GOOGLE_SHEETS_CREDENTIALS_JSON not found in environment")
            return False
            
        # Parse the credentials JSON and fix the private key
        try:
            credentials_info = json.loads(creds_json)
            # Fix the private key by replacing literal \n with actual newlines
            if 'private_key' in credentials_info:
                credentials_info['private_key'] = credentials_info['private_key'].replace('\\n', '\n')
        except json.JSONDecodeError as e:
            print(f"ERROR: Failed to parse credentials JSON: {e}")
            return False
        
        # Create credentials object from the service account info
        creds = ServiceAccountCredentials.from_service_account_info(
            credentials_info, scopes=SCOPES
        )
        
        print("SUCCESS: Successfully authenticated with Google Sheets API")
        
        # Build the service
        service = build('sheets', 'v4', credentials=creds)
        
        # Get the spreadsheet ID from environment
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID')
        print(f"Using spreadsheet ID: {SPREADSHEET_ID}")
        
        if not SPREADSHEET_ID:
            print("ERROR: GOOGLE_SHEET_ID not found in environment")
            return False
        
        # Get spreadsheet info
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        print(f"SUCCESS: Successfully accessed spreadsheet: {spreadsheet['properties']['title']}")
        
        # List the sheets in the spreadsheet
        sheets = spreadsheet['sheets']
        print("\nSheets in the spreadsheet:")
        for i, sheet in enumerate(sheets):
            sheet_title = sheet['properties']['title']
            sheet_id = sheet['properties']['sheetId']
            print(f"  {i+1}. '{sheet_title}' (ID: {sheet_id})")
        
        print("\nThe application expects these specific sheet names:")
        print("  - 'Quizzes'")
        print("  - 'DetailedQuizAnswers'")
        print("  - 'Attendance' (for attendance records)")
        print("  - 'Assignments' (for assignment submissions)")
        print("\nPlease create these sheets in your Google Spreadsheet if they don't exist.")
        
        return True
        
    except Exception as e:
        print(f"ERROR getting spreadsheet info: {e}")
        return False

if __name__ == "__main__":
    list_spreadsheet_sheets()