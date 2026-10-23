import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import psycopg2
from psycopg2.extras import execute_values
from datetime import datetime

creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
sheet_id = '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'
database_url = 'postgresql://postgres.zkewrdlwgtxselynkrmf:Tw%236%2F5rZ37hri%2BK@aws-0-ap-northeast-1.pooler.supabase.com:5432/postgres'

info = json.loads(creds_json)
creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
service = build('sheets', 'v4', credentials=creds)

conn = psycopg2.connect(database_url)
cursor = conn.cursor()

print("🔧 Fixing student created_at fields in Supabase...")
cursor.execute("UPDATE student SET created_at = NOW() WHERE created_at IS NULL;")
conn.commit()
print("✅ Students updated.")

print("🗑️ Clearing existing excel_submission records...")
cursor.execute("DELETE FROM excel_submission;")
conn.commit()
print("✅ excel_submission cleared.")

print("🚀 Fetching exact range A1:H383 from 'Excel Assignments' sheet...")
result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'Excel Assignments'!A1:H383").execute()
rows = result.get('values', [])

if not rows or len(rows) <= 1:
    print("No rows found.")
    exit()

header = [h.strip().lower() for h in rows[0]]
print(f"Header: {header}")
print(f"Total rows fetched: {len(rows)-1}")

submissions_to_add = []
students_to_add = set()

for r in rows[1:]:
    while len(r) < 8:
        r.append('')
        
    student_id = r[0].strip()
    name = r[1].strip()
    assignment_title = r[2].strip()
    score_raw = r[3].strip()
    percentage_raw = r[4].strip()
    submitted_at_str = r[5].strip()
    sync_time_str = r[6].strip()
    grade_details = r[7].strip() # Column H AI response
    
    if not student_id:
        continue
        
    students_to_add.add((student_id, name if name else 'Student'))
    
    # Parse score
    score_val = 0.0
    if '/' in score_raw:
        try:
            score_val = float(score_raw.split('/')[0])
        except:
            score_val = 0.0
    else:
        try:
            score_val = float(score_raw)
        except:
            score_val = 0.0
            
    # Parse percentage
    pct_val = 0.0
    if '%' in percentage_raw:
        try:
            pct_val = float(percentage_raw.replace('%', '').strip())
        except:
            pct_val = 0.0
    else:
        try:
            pct_val = float(percentage_raw)
        except:
            pct_val = 0.0

    assignment_id = 1
    if 'Excel Skill 2' in assignment_title:
        assignment_id = 2
    elif 'Excel Skill 3' in assignment_title:
        assignment_id = 3
    elif 'Excel Skill 4' in assignment_title:
        assignment_id = 4
    elif 'Excel Skill 5' in assignment_title:
        assignment_id = 5
    elif 'Excel Skill 6' in assignment_title:
        assignment_id = 6

    submissions_to_add.append((
        assignment_id,
        student_id,
        score_val,
        pct_val,
        grade_details if grade_details else 'CLEAN',
        'submitted'
    ))

print(f"Ensuring {len(students_to_add)} students exist...")
execute_values(
    cursor,
    "INSERT INTO student (student_id, name, password_hash, created_at) VALUES %s ON CONFLICT (student_id) DO UPDATE SET name = EXCLUDED.name;",
    [(s[0], s[1], 'pbkdf2:sha256:600000$dummyhash', datetime.now()) for s in students_to_add]
)
conn.commit()

print(f"Inserting {len(submissions_to_add)} Excel submissions with Column H grade_details...")
execute_values(
    cursor,
    """
    INSERT INTO excel_submission (assignment_id, student_id, score, percentage, grade_details, status)
    VALUES %s
    """,
    submissions_to_add
)
conn.commit()

print(f"✅ Successfully migrated {len(submissions_to_add)} Excel submissions with AI responses into Supabase!")
cursor.close()
conn.close()
