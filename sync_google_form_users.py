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
SPREADSHEET_ID = '1N23HvM_BvBEKDVi-q1m76ZozEmk1MSbPMitW6rltK5c'
SHEET_NAME = 'Form Responses 1'
ROLL_NUMBER_COL_INDEX = 7  # Column H is 0-indexed index 7

from clean_sheets_sync import get_sheets_service

def sync_users_from_sheet():
    print(f"🚀 Starting User Sync from Google Sheet: {SPREADSHEET_ID}")
    
    # Use the existing service getter from clean_sheets_sync
    service, _ = get_sheets_service()
    if not service:
        print("❌ Could not initialize Google Sheets service. Check environment variables.")
        return

    try:
        # Fetch data from the sheet
        # We use the specific SPREADSHEET_ID provided by the user
        range_name = f"'{SHEET_NAME}'!A:Z"
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
        print(f"Headers: {headers}")
        
        # Check if Column H exists
        if len(headers) <= ROLL_NUMBER_COL_INDEX:
            print(f"❌ Column H (Roll Number) not found. Total columns: {len(headers)}")
            return

        new_users_count = 0
        existing_users_count = 0
        
        with app.app_context():
            for i, row in enumerate(values[1:], start=2):
                if len(row) <= ROLL_NUMBER_COL_INDEX:
                    continue
                
                roll_number = str(row[ROLL_NUMBER_COL_INDEX]).strip()
                if not roll_number or roll_number.lower() == 'roll number':
                    continue
                
                # Try to get Name (assuming it's in Column B or similar)
                # If we don't know the name column, we use roll_number as name for now
                name = roll_number
                if len(row) > 1: # Usually Name is second column
                    name = str(row[1]).strip()
                
                # Check if student already exists
                student = Student.query.filter_by(student_id=roll_number).first()
                
                if not student:
                    print(f"➕ Creating new user: {roll_number}")
                    new_student = Student(
                        student_id=roll_number,
                        name=name
                    )
                    # Username and Password are the same (roll_number)
                    new_student.set_password(roll_number)
                    db.session.add(new_student)
                    new_users_count += 1
                else:
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

if __name__ == "__main__":
    sync_users_from_sheet()
