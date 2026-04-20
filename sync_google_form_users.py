import os
import json
import sys
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from werkzeug.security import generate_password_hash

# Add the app directory to the path to import models
sys.path.insert(0, '.')
from app import app, db, Student

# Google Sheet Configuration
SPREADSHEET_ID = os.environ.get('GOOGLE_SHEET_ID', '1N23HvM_BvBEKDVi-q1m76ZozEmk1MSbPMitW6rltK5c')
SHEET_NAME = 'username'  # User requested sheet name 'username'
# Fallback sheet names if 'username' is not found
FALLBACK_SHEET_NAMES = ['users', 'Form Responses 1', 'Sheet1']

from clean_sheets_sync import get_sheets_service

def sync_users_from_sheet():
    print(f"🚀 Starting User Sync from Google Sheet: {SPREADSHEET_ID}")
    
    # Use the existing service getter from clean_sheets_sync
    service, _ = get_sheets_service()
    if not service:
        print("❌ Could not initialize Google Sheets service. Check environment variables.")
        return

    try:
        # First, try to find which sheet exists
        spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
        
        target_sheet = None
        if SHEET_NAME in sheets:
            target_sheet = SHEET_NAME
        else:
            for fallback in FALLBACK_SHEET_NAMES:
                if fallback in sheets:
                    target_sheet = fallback
                    break
        
        if not target_sheet:
            print(f"❌ Could not find target sheet '{SHEET_NAME}' or any fallbacks in {sheets}")
            return

        print(f"Using sheet: {target_sheet}")
        
        # Fetch data from the sheet
        range_name = f"'{target_sheet}'!A:Z"
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print("No data found in the sheet.")
            return

        # Header is values[0], data starts from values[1]
        headers = values[0]
        print(f"Headers found: {headers}")
        
        # Dynamically find column indices
        student_id_col = -1
        name_col = -1
        
        # Keywords to look for
        student_id_keywords = ['studentid', 'roll number', 'student id', 'roll_number', 'user id', 'username']
        name_keywords = ['login name', 'name', 'full name', 'student name', 'login_name']
        
        for i, header in enumerate(headers):
            header_lower = str(header).lower().strip()
            if student_id_col == -1 and any(k in header_lower for k in student_id_keywords):
                student_id_col = i
            elif name_col == -1 and any(k in header_lower for k in name_keywords):
                name_col = i

        # If not found, use defaults
        if student_id_col == -1:
            student_id_col = 0 # Default to first column
            print(f"⚠️ StudentID column not found by header, using column index {student_id_col}")
        if name_col == -1:
            name_col = 1 if len(headers) > 1 else 0
            print(f"⚠️ Name column not found by header, using column index {name_col}")

        print(f"Using columns: StudentID at {student_id_col}, Name at {name_col}")

        new_users_count = 0
        existing_users_count = 0
        
        with app.app_context():
            for i, row in enumerate(values[1:], start=2):
                if len(row) <= max(student_id_col, name_col):
                    # Pad row if it's shorter than expected columns
                    row.extend([''] * (max(student_id_col, name_col) - len(row) + 1))
                
                student_id = str(row[student_id_col]).strip()
                if not student_id or student_id.lower() in ['studentid', 'roll number', 'student id']:
                    continue
                
                name = str(row[name_col]).strip() if name_col < len(row) else student_id
                if not name:
                    name = student_id
                
                # Check if student already exists
                student = Student.query.filter_by(student_id=student_id).first()
                
                if not student:
                    print(f"➕ Creating new user: {student_id} ({name})")
                    new_student = Student(
                        student_id=student_id,
                        name=name
                    )
                    # Username and Password are the same (student_id)
                    new_student.set_password(student_id)
                    db.session.add(new_student)
                    new_users_count += 1
                else:
                    # Update name if it's different and not empty? 
                    # For now, just count existing
                    existing_users_count += 1
            
            if new_users_count > 0:
                db.session.commit()
                print(f"✅ Successfully created {new_users_count} new users.")
            else:
                print("ℹ️ No new users to create.")
            
            print(f"📊 Total processed: {new_users_count + existing_users_count}")
            print(f"📊 Existing users: {existing_users_count}")

    except Exception as e:
        print(f"❌ Error during sync: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    sync_users_from_sheet()
