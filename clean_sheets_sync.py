"""
Clean Google Sheets Sync - ALL DATA TYPES
Proper columns, proper sheets, proper sync!
"""
import os
import json
import requests
from datetime import datetime

# Cache for service
_google_service = None
_google_sheet_id = None

SHEET_HEADERS = {
    'Students': ['Student ID', 'Name', 'Created At', 'Sync Time'],
    'Attendance': ['Date', 'Student ID', 'Name', 'Check-In Time', 'Check-Out Time', 'Status', 
                   'Check-In Location (Lat,Lng)', 'Check-In Latitude', 'Check-In Longitude', 'Check-In Address',
                   'Check-Out Location (Lat,Lng)', 'Check-Out Latitude', 'Check-Out Longitude', 'Check-Out Address',
                   'Sync Timestamp'],
    'Quiz Results': ['Student ID', 'Name', 'Quiz Title', 'Score', 'Percentage', 'Submitted At', 'Sync Time'],
    'Assignments': ['Student ID', 'Name', 'Assignment Title', 'Submission URL', 'Grade', 'Submitted At', 'Sync Time'],
    'Midterm Grades': ['Student ID', 'Name', 'Midterm Title', 'Grade', 'Graded At', 'Sync Time'],
    'Course Outlines': ['Course Title', 'Created At', 'Sync Time']
}

def get_address_from_coordinates(lat, lng):
    """Convert GPS coordinates to human-readable address using OpenStreetMap API"""
    try:
        url = f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lng}&addressdetails=1"
        headers = {'User-Agent': 'ERP-System-Attendance-Tracker/1.0'}
        response = requests.get(url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            return data.get('display_name', f"Coordinates: {lat}, {lng}")
        return f"Coordinates: {lat}, {lng}"
    except Exception as e:
        print(f"⚠️ Reverse geocoding error: {e}")
        return f"Coordinates: {lat}, {lng}"

def get_sheets_service():
    """Get Google Sheets service"""
    global _google_service, _google_sheet_id
    
    if _google_service:
        return _google_service, _google_sheet_id
    
    creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    
    if not creds_json or not sheet_id:
        print("⚠️ Google Sheets not configured")
        return None, None
    
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        
        info = json.loads(creds_json)
        if "private_key" in info:
            info["private_key"] = info["private_key"].replace("\\n", "\n")
        
        creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets'])
        service = build('sheets', 'v4', credentials=creds)
        
        _google_service = service
        _google_sheet_id = sheet_id
        
        print("✅ Google Sheets service ready!")
        return service, sheet_id
    except Exception as e:
        print(f"❌ Google Sheets init failed: {e}")
        return None, None

def ensure_sheet_exists(service, sheet_id, sheet_name):
    """Create sheet if not exists and add headers"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
        
        if sheet_name not in sheets:
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
            ).execute()
            
            # Add headers
            headers = SHEET_HEADERS.get(sheet_name, [])
            if headers:
                service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f'{sheet_name}!A1',
                    valueInputOption='USER_ENTERED',
                    body={'values': [headers]}
                ).execute()
            print(f"✅ Created sheet: {sheet_name}")
        return True
    except Exception as e:
        print(f"❌ Error creating sheet {sheet_name}: {e}")
        return False

def sync_to_sheets(sheet_name, data_row):
    """Generic sync function - adds a row to any sheet"""
    try:
        service, sheet_id = get_sheets_service()
        if not service:
            return False
        
        # Ensure sheet exists with headers
        ensure_sheet_exists(service, sheet_id, sheet_name)
        
        # Add sync timestamp
        data_row.append(str(datetime.now()))
        
        # Append row
        service.spreadsheets().values().append(
            spreadsheetId=sheet_id,
            range=f'{sheet_name}!A2',
            valueInputOption='USER_ENTERED',
            body={'values': [data_row]}
        ).execute()
        
        print(f"✅ Synced to {sheet_name}: {data_row[0]}")
        return True
    except Exception as e:
        print(f"❌ Sync to {sheet_name} FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False

# ============================================
# SPECIFIC SYNC FUNCTIONS
# ============================================

def sync_student(student_id, name):
    """Sync student to Google Sheets"""
    return sync_to_sheets('Students', [
        student_id,
        name,
        str(datetime.now())
    ])

def sync_attendance(student_id, name, date, check_in, check_out, status, check_in_location=None, check_out_location=None, address=None):
    """Sync attendance to Google Sheets with ALL 15 columns"""
    try:
        service, sheet_id = get_sheets_service()
        if not service:
            return False
        
        # Ensure sheet exists with proper headers
        attendance_headers = [
            'Date', 'Student ID', 'Name', 'Check-In Time', 'Check-Out Time', 'Status',
            'Check-In Location (Lat,Lng)', 'Check-In Latitude', 'Check-In Longitude', 'Check-In Address',
            'Check-Out Location (Lat,Lng)', 'Check-Out Latitude', 'Check-Out Longitude', 'Check-Out Address',
            'Sync Timestamp'
        ]
        
        # Ensure sheet exists
        try:
            spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
            sheets = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
            
            if 'Attendance' not in sheets:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id,
                    body={'requests': [{'addSheet': {'properties': {'title': 'Attendance'}}}]}
                ).execute()
                # Add headers
                service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range='Attendance!A1',
                    valueInputOption='USER_ENTERED',
                    body={'values': [attendance_headers]}
                ).execute()
                print("✅ Created Attendance sheet with headers")
        except Exception as e:
            print(f"⚠️ Error setting up sheet: {e}")
        
        # Parse check-in location
        check_in_lat = ''
        check_in_lng = ''
        check_in_addr = address or ''
        
        if check_in_location:
            try:
                parts = str(check_in_location).split(',')
                if len(parts) == 2:
                    check_in_lat = parts[0].strip()
                    check_in_lng = parts[1].strip()
                    # Get address if not provided
                    if not check_in_addr:
                        check_in_addr = get_address_from_coordinates(check_in_lat, check_in_lng)
            except Exception as e:
                print(f"⚠️ Error parsing check-in location: {e}")
        
        # Parse check-out location
        check_out_lat = ''
        check_out_lng = ''
        check_out_addr = ''
        
        if check_out_location:
            try:
                parts = str(check_out_location).split(',')
                if len(parts) == 2:
                    check_out_lat = parts[0].strip()
                    check_out_lng = parts[1].strip()
                    # Get address for check-out too
                    check_out_addr = get_address_from_coordinates(check_out_lat, check_out_lng)
            except Exception as e:
                print(f"⚠️ Error parsing check-out location: {e}")
        
        # Build row with ALL 14 data columns (15th is sync timestamp added by sync_to_sheets)
        data_row = [
            str(date),                              # Date
            student_id,                             # Student ID
            name,                                   # Name
            str(check_in) if check_in else '',      # Check-In Time
            str(check_out) if check_out else '',    # Check-Out Time
            status,                                 # Status
            check_in_location or '',                # Check-In Location (Lat,Lng)
            check_in_lat,                           # Check-In Latitude
            check_in_lng,                           # Check-In Longitude
            check_in_addr,                          # Check-In Address
            check_out_location or '',               # Check-Out Location (Lat,Lng)
            check_out_lat,                          # Check-Out Latitude
            check_out_lng,                          # Check-Out Longitude
            check_out_addr                          # Check-Out Address
        ]
        
        return sync_to_sheets('Attendance', data_row)
    except Exception as e:
        print(f"❌ Attendance sync FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False

def sync_quiz(student_id, name, quiz_title, score, total_questions, submitted_at):
    """Sync quiz result to Google Sheets"""
    percentage = f"{(score/total_questions*100):.1f}%" if total_questions > 0 else "0%"
    return sync_to_sheets('Quiz Results', [
        student_id,
        name,
        quiz_title,
        f"{score}/{total_questions}",
        percentage,
        str(submitted_at),
        str(datetime.now())
    ])

def sync_assignment(student_id, name, assignment_title, submission_url, grade, submitted_at):
    """Sync assignment to Google Sheets"""
    return sync_to_sheets('Assignments', [
        student_id,
        name,
        assignment_title,
        submission_url,
        str(grade) if grade else '',
        str(submitted_at),
        str(datetime.now())
    ])

def sync_midterm(student_id, name, midterm_title, grade, graded_at):
    """Sync midterm grade to Google Sheets"""
    return sync_to_sheets('Midterm Grades', [
        student_id,
        name,
        midterm_title,
        grade,
        str(graded_at),
        str(datetime.now())
    ])
