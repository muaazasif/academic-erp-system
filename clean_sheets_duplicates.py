import os
import json
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

def get_sheets_service():
    """Initialize Google Sheets service"""
    try:
        creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
        sheet_id = os.environ.get('GOOGLE_SHEET_ID')
        
        if not creds_json or not sheet_id:
            print("❌ GOOGLE_CREDENTIALS_JSON or GOOGLE_SHEET_ID not found in environment")
            return None, None
            
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
        service = build('sheets', 'v4', credentials=creds)
        return service, sheet_id
    except Exception as e:
        print(f"❌ Error initializing service: {e}")
        return None, None

def cleanup_duplicates():
    service, sheet_id = get_sheets_service()
    if not service:
        return

    sheet_name = 'Excel Assignments'
    print(f"🔍 Fetching data from '{sheet_name}'...")

    try:
        # Get all data
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A:H"
        ).execute()
        
        rows = result.get('values', [])
        if not rows:
            print("No data found.")
            return

        header = rows[0]
        data_rows = rows[1:]
        
        # Dictionary to store the best/latest submission for each student+assignment
        # Key: (student_id, assignment_name)
        # Value: Row data
        unique_entries = {}

        print(f"Processing {len(data_rows)} rows...")

        for row in data_rows:
            if len(row) < 3: continue
            
            student_id = str(row[0]).strip()
            assignment = str(row[2]).strip()
            key = (student_id, assignment)
            
            # If we haven't seen this student+assignment yet, or if this one is better
            # For now, let's keep the LATEST one (based on index in the list)
            # Since rows are usually appended, the last one in the list is the latest.
            unique_entries[key] = row

        # Prepare new values starting with header
        new_values = [header]
        
        # Sort by Student ID or just add them back
        # We'll sort by Assignment then Student ID for better organization
        sorted_keys = sorted(unique_entries.keys(), key=lambda x: (x[1], x[0]))
        for key in sorted_keys:
            new_values.append(unique_entries[key])

        print(f"✅ Found {len(unique_entries)} unique submissions (removed {len(data_rows) - len(unique_entries)} duplicates).")

        # Clear the sheet first
        service.spreadsheets().values().clear(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A:Z"
        ).execute()

        # Write unique data back
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': new_values}
        ).execute()

        print("🚀 Google Sheet successfully cleaned and deduplicated!")

    except Exception as e:
        print(f"❌ Cleanup failed: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()
    cleanup_duplicates()
