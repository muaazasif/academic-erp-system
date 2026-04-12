"""
Export ALL Database Data to Google Sheets
This uploads everything: Students, Attendance, Quizzes, Assignments, etc.
"""
import os
import sqlite3
from datetime import datetime
import json

# Database path
DB_PATH = 'instance/erp_system.db'

def get_google_credentials():
    """Load Google Sheets credentials"""
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        
        creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
        
        if not creds_json:
            # Try loading from file
            if os.path.exists('credentials.json'):
                with open('credentials.json', 'r') as f:
                    creds_json = f.read()
            else:
                print("❌ No Google credentials found!")
                return None, None
        
        creds_dict = json.loads(creds_json)
        
        SCOPES = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        
        return service, creds
    except Exception as e:
        print(f"❌ Error loading credentials: {e}")
        return None, None

def create_or_get_sheet(service, spreadsheet_id, sheet_name):
    """Create a new sheet or get existing one"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        # Check if sheet already exists
        for sheet in spreadsheet.get('sheets', []):
            if sheet['properties']['title'] == sheet_name:
                # Clear existing data
                service.spreadsheets().values().clear(
                    spreadsheetId=spreadsheet_id,
                    range=sheet_name
                ).execute()
                return sheet_name
        
        # Create new sheet
        request_body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name
                    }
                }
            }]
        }
        
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=request_body
        ).execute()
        
        return sheet_name
    except Exception as e:
        print(f"❌ Error creating sheet {sheet_name}: {e}")
        return None

def upload_data_to_sheet(service, spreadsheet_id, sheet_name, headers, data):
    """Upload data to Google Sheet"""
    try:
        # Prepare data for upload
        values = [headers]  # Header row
        
        for row in data:
            # Convert row to list in correct order
            row_data = []
            for header in headers:
                value = row.get(header, '')
                # Convert None to empty string
                if value is None:
                    value = ''
                # Convert datetime to string
                elif isinstance(value, (datetime,)):
                    value = str(value)
                row_data.append(str(value))
            values.append(row_data)
        
        # Upload to Google Sheets
        body = {
            'values': values
        }
        
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!A1',
            valueInputOption='RAW',
            body=body
        ).execute()
        
        print(f"  ✅ Uploaded {len(data)} rows to '{sheet_name}'")
        return True
        
    except Exception as e:
        print(f"  ❌ Error uploading to {sheet_name}: {e}")
        return False

def export_all_to_sheets():
    """Main export function"""
    print("=" * 70)
    print("📤 EXPORTING ALL DATA TO GOOGLE SHEETS")
    print("=" * 70)
    
    # Get credentials
    service, creds = get_google_credentials()
    if not service:
        print("\n❌ Failed to load Google credentials!")
        print("💡 Make sure GOOGLE_SHEETS_CREDENTIALS_JSON is set in your environment")
        return False
    
    # Get spreadsheet ID
    SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID')
    if not SPREADSHEET_ID:
        print("\n❌ No spreadsheet ID found!")
        print("💡 Set GOOGLE_SHEET_ID environment variable")
        return False
    
    print(f"\n✅ Connected to Google Sheets")
    print(f"📋 Spreadsheet ID: {SPREADSHEET_ID}")
    
    # Connect to SQLite
    if not os.path.exists(DB_PATH):
        print(f"\n❌ Database not found at {DB_PATH}")
        return False
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Get all tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [row['name'] for row in cursor.fetchall()]
    
    print(f"\n📦 Found {len(tables)} tables to export")
    print()
    
    success_count = 0
    
    for table in tables:
        try:
            print(f"📊 Exporting: {table}")
            
            # Get all data from table
            cursor.execute(f"SELECT * FROM {table}")
            rows = cursor.fetchall()
            
            if not rows:
                print(f"  ℹ️ Table is empty, skipping")
                continue
            
            # Get column names
            headers = list(rows[0].keys())
            
            # Convert rows to dictionaries
            data = [dict(row) for row in rows]
            
            # Create or clear sheet
            sheet_name = table.replace('_', ' ').title()
            create_or_get_sheet(service, SPREADSHEET_ID, sheet_name)
            
            # Upload data
            if upload_data_to_sheet(service, SPREADSHEET_ID, sheet_name, headers, data):
                success_count += 1
            
        except Exception as e:
            print(f"  ❌ Error exporting {table}: {e}")
    
    conn.close()
    
    # Summary
    print("\n" + "=" * 70)
    print(f"✅ EXPORT COMPLETE!")
    print(f"📊 Successfully exported {success_count}/{len(tables)} tables")
    print(f"🔗 View your data: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print("=" * 70)
    
    return True

if __name__ == '__main__':
    export_all_to_sheets()
