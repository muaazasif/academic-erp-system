import os
import json
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build

# Load environment variables
load_dotenv()

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def ensure_sheets_exist():
    """Ensure required sheets exist in the spreadsheet"""
    try:
        print("Ensuring required sheets exist in the spreadsheet...")
        
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
        
        # Get current spreadsheet info
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        existing_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
        
        print(f"Current sheets: {existing_sheets}")
        
        # Define required sheets
        required_sheets = ['Quizzes', 'DetailedQuizAnswers', 'Attendance', 'Assignments']
        
        # Find the highest sheet ID to use as a basis for new sheets
        max_sheet_id = max(sheet['properties']['sheetId'] for sheet in spreadsheet['sheets'])
        
        # Check for missing sheets and create them
        sheets_to_create = []
        for req_sheet in required_sheets:
            if req_sheet not in existing_sheets:
                sheets_to_create.append({
                    'addSheet': {
                        'properties': {
                            'title': req_sheet,
                            'sheetId': max_sheet_id + len(sheets_to_create) + 1
                        }
                    }
                })
        
        if sheets_to_create:
            print(f"Creating missing sheets: {[s['addSheet']['properties']['title'] for s in sheets_to_create]}")
            
            # Batch update to add the missing sheets
            body = {
                'requests': sheets_to_create
            }
            
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=body
            ).execute()
            
            print(f"Successfully created {len(sheets_to_create)} sheets")
        else:
            print("All required sheets already exist!")
        
        # Add headers to each sheet if they don't exist
        for sheet_name in required_sheets:
            # Check if headers already exist by reading first row
            result = service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{sheet_name}'!1:1"
            ).execute()
            
            values = result.get('values', [])
            if not values or len(values[0]) == 0:
                # Add appropriate headers based on sheet type
                if sheet_name == 'Quizzes':
                    headers = ['Student ID', 'Name', 'Quiz Title', 'Score', 'Total Questions', 'Percentage', 'Submitted At', 'Synced At']
                elif sheet_name == 'DetailedQuizAnswers':
                    headers = ['Student ID', 'Name', 'Quiz Title', 'Question #', 'Question Text', 'Selected Option', 'Correct Option', 'Result', 'Submitted At', 'Synced At']
                elif sheet_name == 'Attendance':
                    headers = ['Date', 'Student ID', 'Name', 'Check In', 'Check Out', 'Status', 'Synced At']
                elif sheet_name == 'Assignments':
                    headers = ['Student ID', 'Name', 'Assignment Title', 'Submission URL', 'Submitted At', 'Grade', 'Synced At']
                
                # Add headers to the sheet
                body = {
                    'values': [headers]
                }
                
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=f"'{sheet_name}'!A1",
                    valueInputOption='RAW',
                    body=body
                ).execute()
                
                print(f"Added headers to {sheet_name} sheet")
            else:
                print(f"Headers already exist in {sheet_name} sheet")
        
        return True
        
    except Exception as e:
        print(f"ERROR ensuring sheets exist: {e}")
        return False

if __name__ == "__main__":
    ensure_sheets_exist()