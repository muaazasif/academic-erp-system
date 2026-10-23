import os
import json
import time
from datetime import datetime
import pytz

# Allow HTTP transport for local development (Railway uses HTTPS automatically)
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    'https://www.googleapis.com/auth/classroom.announcements',
    'https://www.googleapis.com/auth/youtube.force-ssl'
]
COURSE_ID = 'ODg0Njk2MjE4NTcy'

CREDENTIALS_DATA = {
    "web": {
        "client_id": os.getenv("GOOGLE_CLIENT_ID", ""),
        "project_id": "forward-venture-484907-c8",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_secret": os.getenv("GOOGLE_CLIENT_SECRET", ""),
        "redirect_uris": ["http://127.0.0.1:5000/oauth2callback"]
    }
}

TOKEN_FILE = 'classroom_token.json'

def get_classroom_service():
    """Authenticate and return Google Classroom service"""
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE)
            if creds and creds.expired and creds.refresh_token:
                from google.auth.transport.requests import Request
                creds.refresh(Request())
        except Exception as e:
            print(f"⚠️ Error loading token file: {e}")
            creds = None
            
    if creds and creds.valid:
        return build('classroom', 'v1', credentials=creds)
    print("⚠️ Classroom credentials not valid or missing token file.")
    return None

def calculate_attendance_stats(session_type='all'):
    """Calculate attendance stats for Bano Qabil 5.0 Fall batch"""
    from app import app, db, Student, Attendance, Assessment
    with app.app_context():
        today = datetime.today().date()
        
        # Get BQ 5.0 Fall active students
        student_query = Student.query.filter_by(batch='Bano Qabil 5.0 Fall', is_active=True)
        
        if session_type == 'girls':
            # Girls timing 10 to 12
            student_query = student_query.filter(Student.timings.ilike('%10 to 12%'))
        elif session_type == 'boys':
            # Boys timing 12 to 2
            student_query = student_query.filter(Student.timings.ilike('%12 to 2%'))
            
        students = student_query.all()
        total_students = len(students)
        
        if total_students == 0:
            return 0, 0, 0, 0.0
            
        student_ids = [s.student_id for s in students]
        
        present_count = Attendance.query.filter(
            Attendance.date == today,
            Attendance.status == 'present',
            Attendance.student_id.in_(student_ids)
        ).count()
        
        absent_count = total_students - present_count
        if absent_count < 0: absent_count = 0
        
        attendance_rate = round((present_count / total_students) * 100, 1) if total_students > 0 else 0.0
        
        return total_students, present_count, absent_count, attendance_rate

def post_attendance_announcement(session_type='all'):
    """Post Daily Attendance Report to Google Classroom"""
    total, present, absent, rate = calculate_attendance_stats(session_type)
    
    session_label = ""
    if session_type == 'girls':
        session_label = " (Girls Session - 10:00 AM)"
    elif session_type == 'boys':
        session_label = " (Boys Session - 12:00 PM)"

    message = f"""📊 Daily Attendance Report{session_label}
Assalam-o-Alaikum Students!

Today's attendance report:
- Total Students: {total}
- Present: {present}
- Absent: {absent}
- Attendance: {rate}%
Please contact the instructor if you believe your attendance record is incorrect."""

    print(f"📢 Generating Attendance Announcement:\n{message}")
    
    service = get_classroom_service()
    if service:
        try:
            announcement = {
                'text': message,
                'state': 'PUBLISHED'
            }
            result = service.courses().announcements().create(courseId=COURSE_ID, body=announcement).execute()
            print(f"✅ Successfully posted attendance announcement to Google Classroom! ID: {result.get('id')}")
            return True, "Posted successfully"
        except Exception as e:
            print(f"❌ Error posting to Google Classroom: {e}")
            return False, str(e)
    else:
        print("⚠️ Google Classroom service not authenticated. Announcement logged only.")
        return False, "Not authenticated"

def run_classroom_attendance_scheduler():
    """Background scheduler checking Saturday and Sunday every 30 minutes"""
    import threading
    def scheduler_loop():
        print("🚀 Google Classroom Attendance Scheduler started (Sat/Sun every 30 mins)...")
        while True:
            try:
                pakistan_tz = pytz.timezone('Asia/Karachi')
                now = datetime.now(pakistan_tz)
                weekday = now.weekday() # 5 = Saturday, 6 = Sunday
                
                # Check if Saturday (5) or Sunday (6)
                if weekday in [5, 6]:
                    hour = now.hour
                    minute = now.minute
                    
                    # Girls timing check: 10:00 AM to 11:59 AM
                    if hour == 10 or hour == 11:
                        # Post once around 10:00 or 10:30 if not already posted
                        # For simplicity, we can trigger during the window
                        pass
                    # Boys timing check: 12:00 PM to 2:00 PM
                    elif hour == 12 or (hour == 13 and minute <= 59):
                        pass
            except Exception as e:
                print(f"⚠️ Classroom scheduler error: {e}")
            time.sleep(1800) # Check every 30 minutes
            
    t = threading.Thread(target=scheduler_loop, daemon=True)
    t.start()
