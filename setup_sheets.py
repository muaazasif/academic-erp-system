import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Path to the external credentials file
CREDENTIALS_PATH = r'E:\Governor Sindh Course\Application\myapp\gen-lang-client-0321830476-1f4575af7b34.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def setup_google_sheets():
    """Setup required sheets in Google Spreadsheet"""
    try:
        print("Setting up Google Sheets...")
        
        # Authenticate using the service account
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        
        # Your spreadsheet ID
        SPREADSHEET_ID = '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'
        
        # Get current spreadsheet info
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        existing_sheets = {sheet['properties']['title']: sheet['properties']['sheetId'] 
                          for sheet in spreadsheet['sheets']}
        
        print(f"Existing sheets: {list(existing_sheets.keys())}")
        
        # Required sheets
        required_sheets = ['Quizzes', 'DetailedQuizAnswers', 'Attendance', 'Assignments']
        
        # Find the highest sheet ID to use as a basis for new sheets
        max_sheet_id = max(sheet['properties']['sheetId'] for sheet in spreadsheet['sheets'])
        
        # Create requests to add missing sheets
        requests = []
        for sheet_name in required_sheets:
            if sheet_name not in existing_sheets:
                requests.append({
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                            'sheetId': max_sheet_id + len(requests) + 1,
                            'gridProperties': {
                                'rowCount': 1000,
                                'columnCount': 20
                            }
                        }
                    }
                })
        
        if requests:
            print(f"Creating sheets: {[req['addSheet']['properties']['title'] for req in requests]}")
            
            # Execute the batch update
            body = {'requests': requests}
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=body
            ).execute()
            
            print(f"Successfully created {len(requests)} sheets")
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
        
        print("Google Sheets setup completed successfully!")
        return True
        
    except Exception as e:
        print(f"Error setting up Google Sheets: {e}")
        return False

if __name__ == "__main__":
    setup_google_sheets()