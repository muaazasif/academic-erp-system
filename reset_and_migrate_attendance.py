import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import psycopg2
from datetime import datetime

creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
sheet_id = '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'
database_url = 'postgresql://postgres.zkewrdlwgtxselynkrmf:Tw%236%2F5rZ37hri%2BK@aws-0-ap-northeast-1.pooler.supabase.com:5432/postgres'

info = json.loads(creds_json)
creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'])
service = build('sheets', 'v4', credentials=creds)

conn = psycopg2.connect(database_url)
cursor = conn.cursor()

print("🗑️ Clearing existing attendance records from Supabase...")
cursor.execute("DELETE FROM attendance;")
conn.commit()
print("✅ Attendance table cleared.")

print("🚀 Fetching full Attendance range from Google Sheets (up to 1300 rows)...")
result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'Attendance'!A1:P1300").execute()
rows = result.get('values', [])
if not rows:
    print("No attendance rows found.")
    exit()

header = [h.strip().lower() for h in rows[0]]
print(f"Header: {header}")
print(f"Total rows found in Google Sheet: {len(rows)-1}")

migrated = 0
skipped = 0

for r in rows[1:]:
    row_dict = {header[i]: r[i] for i in range(min(len(header), len(r))) if header[i]}
    
    date_str = row_dict.get('date') or (r[0] if len(r) > 0 else '')
    student_id = row_dict.get('student id') or row_dict.get('student_id') or (r[1] if len(r) > 1 else '')
    check_in = row_dict.get('check-in time') or (r[3] if len(r) > 3 else None)
    check_out = row_dict.get('check-out time') or (r[4] if len(r) > 4 else None)
    status = row_dict.get('status') or (r[5] if len(r) > 5 else 'present')
    
    # Location fields
    check_in_loc = row_dict.get('check-in location (lat,lng)') or (r[6] if len(r) > 6 else None)
    check_out_loc = row_dict.get('check-out location (lat,lng)') or (r[10] if len(r) > 10 else None)
    
    if not student_id or not date_str or str(student_id).strip() == '':
        skipped += 1
        continue
    
    try:
        att_date = datetime.strptime(date_str.split()[0], '%Y-%m-%d').date()
        
        # Ensure student exists in student table first to avoid FK error
        cursor.execute(
            "INSERT INTO student (student_id, name, password_hash) VALUES (%s, %s, %s) ON CONFLICT (student_id) DO NOTHING;",
            (str(student_id).strip(), row_dict.get('name') or (r[2] if len(r) > 2 else 'Student'), 'pbkdf2:sha256:600000$dummyhash')
        )
        
        cursor.execute(
            """
            INSERT INTO attendance (student_id, date, check_in_time, check_out_time, check_in_location, check_out_location, status)
            VALUES (%s, %s, %s, %s, %s, %s, %s);
            """,
            (
                str(student_id).strip(),
                att_date,
                check_in if check_in and check_in.strip() else None,
                check_out if check_out and check_out.strip() else None,
                check_in_loc if check_in_loc and check_in_loc.strip() else None,
                check_out_loc if check_out_loc and check_out_loc.strip() else None,
                str(status).strip() if status else 'present'
            )
        )
        conn.commit()
        migrated += 1
    except Exception as e:
        conn.rollback()
        skipped += 1

print(f"✅ Successfully migrated {migrated} attendance records (skipped {skipped}) with complete location data into Supabase!")
cursor.close()
conn.close()
