import os
import json
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build

# Load environment variables
load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def test_google_sheets_connection():
    """Test if Google Sheets connection is working"""
    try:
        print("Testing Google Sheets connection...")

        # Get credentials from environment variable
        creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
        if not creds_json:
            print("ERROR: GOOGLE_SHEETS_CREDENTIALS_JSON not found in environment")
            return False

        # Parse the credentials JSON
        try:
            credentials_info = json.loads(creds_json)
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
        for sheet in sheets:
            print(f"  - {sheet['properties']['title']} (ID: {sheet['properties']['sheetId']})")

        return True

    except Exception as e:
        print(f"ERROR connecting to Google Sheets: {e}")
        return False

if __name__ == "__main__":
    test_google_sheets_connection()