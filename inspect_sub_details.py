import sqlite3
import json

conn = sqlite3.connect('instance/bq-erp.db')
cursor = conn.cursor()
cursor.execute("SELECT id, grade_details FROM excel_submission WHERE id IN (5, 6)")
rows = cursor.fetchall()
if not rows:
    cursor.execute("SELECT id, grade_details FROM excel_submission")
    rows = cursor.fetchall()

for r in rows:
    print(f"ID: {r[0]}")
    parsed = json.loads(r[1])
    print(json.dumps(parsed, indent=2)[:1000])
    print("="*60)
conn.close()
