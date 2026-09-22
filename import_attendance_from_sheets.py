import os
import sys
from datetime import datetime
from app import app, db, Student, Attendance

# Add the app directory to the path to import models
sys.path.insert(0, '.')

from clean_sheets_sync import get_sheets_service

def import_attendance_from_sheet():
    print("🚀 Starting Attendance Import from Google Sheet...")
    
    service, spreadsheet_id = get_sheets_service()
    if not service:
        print("❌ Could not initialize Google Sheets service.")
        return

    # Fetch data from the Attendance sheet
    try:
        range_name = "'Attendance'!A2:F" # Just headers 0-5 for now
        result = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        
        values = result.get('values', [])
        if not values:
            print("No attendance data found in the sheet.")
            return

        print(f"Found {len(values)} records in Google Sheet.")

        with app.app_context():
            new_records = 0
            for row in values:
                # Expected format: Date, Student ID, Name, Check-In, Check-Out, Status
                if len(row) < 6:
                    continue
                    
                date_str = row[0]
                student_id = str(row[1]).strip()
                # name = row[2]
                check_in_str = row[3]
                check_out_str = row[4]
                status = row[5]

                # Convert date
                try:
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                except:
                    continue

                # Parse times
                check_in_time = None
                if check_in_str:
                    try:
                        check_in_time = datetime.strptime(check_in_str, '%H:%M:%S').time()
                    except:
                        pass
                
                check_out_time = None
                if check_out_str:
                    try:
                        check_out_time = datetime.strptime(check_out_str, '%H:%M:%S').time()
                    except:
                        pass

                # Check if record exists
                existing = Attendance.query.filter_by(student_id=student_id, date=date_obj).first()
                if not existing:
                    new_att = Attendance(
                        student_id=student_id,
                        date=date_obj,
                        check_in_time=check_in_time,
                        check_out_time=check_out_time,
                        status=status
                    )
                    db.session.add(new_att)
                    new_records += 1
            
            db.session.commit()
            print(f"✅ Imported {new_records} new attendance records into local database.")

    except Exception as e:
        print(f"❌ Error during import: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import_attendance_from_sheet()
