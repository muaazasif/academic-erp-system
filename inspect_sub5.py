import sqlite3
import json

conn = sqlite3.connect('erp_system.db')
cursor = conn.cursor()

cursor.execute("SELECT id, grade_details FROM excel_submission WHERE id IN (5, 6)")
rows = cursor.fetchall()
for r in rows:
    print(f"ID: {r[0]}")
    print(f"Grade Details Raw: {repr(r[1])}")
    try:
        parsed = json.loads(r[1])
        print(f"Parsed JSON Type: {type(parsed)}")
        print(f"Parsed JSON: {repr(parsed)}")
    except Exception as e:
        print(f"JSON Parse Error: {e}")
    print("="*50)

conn.close()
