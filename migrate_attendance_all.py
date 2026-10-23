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

result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'Attendance'!A1:O1000").execute()
rows = result.get('values', [])
header = [h.strip().lower() for h in rows[0]]
print(f"Found {len(rows)-1} attendance rows to migrate.")

migrated = 0
for r in rows[1:]:
    row_dict = {header[i]: r[i] for i in range(min(len(header), len(r))) if header[i]}
    date_str = row_dict.get('date') or (r[0] if len(r) > 0 else '')
    student_id = row_dict.get('student id') or row_dict.get('student_id') or (r[1] if len(r) > 1 else '')
    status = row_dict.get('status') or (r[5] if len(r) > 5 else 'present')
    
    if student_id and date_str and str(student_id).strip():
        try:
            att_date = datetime.strptime(date_str.split()[0], '%Y-%m-%d').date()
            cursor.execute(
                "INSERT INTO attendance (student_id, date, status) VALUES (%s, %s, %s) ON CONFLICT DO NOTHING;",
                (str(student_id).strip(), att_date, str(status).strip())
            )
            conn.commit()
            migrated += 1
        except Exception:
            conn.rollback()

print(f"✅ Successfully migrated {migrated} attendance records into Supabase!")
cursor.close()
conn.close()
