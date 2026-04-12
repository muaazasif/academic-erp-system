"""
Clean Google Sheets Sync - ALL DATA TYPES
Proper columns, proper sheets, proper sync!
"""
import os
import json
from datetime import datetime

# Cache for service
_google_service = None
_google_sheet_id = None

SHEET_HEADERS = {
    'Students': ['Student ID', 'Name', 'Created At', 'Sync Time'],
    'Attendance': ['Date', 'Student ID', 'Name', 'Check-In Time', 'Check-Out Time', 'Status', 'Check-In Location', 'Check-Out Location', 'Address', 'Sync Time'],
    'Quiz Results': ['Student ID', 'Name', 'Quiz Title', 'Score', 'Percentage', 'Submitted At', 'Sync Time'],
    'Assignments': ['Student ID', 'Name', 'Assignment Title', 'Submission URL', 'Grade', 'Submitted At', 'Sync Time'],
    'Midterm Grades': ['Student ID', 'Name', 'Midterm Title', 'Grade', 'Graded At', 'Sync Time'],
    'Course Outlines': ['Course Title', 'Created At', 'Sync Time']
}

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
    """Sync attendance to Google Sheets"""
    return sync_to_sheets('Attendance', [
        str(date),
        student_id,
        name,
        str(check_in) if check_in else '',
        str(check_out) if check_out else '',
        status,
        check_in_location or '',
        check_out_location or '',
        address or '',
        str(datetime.now())
    ])

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
