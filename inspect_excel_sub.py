import sqlite3

conn = sqlite3.connect('instance/bq-erp.db')
cursor = conn.cursor()
cursor.execute("SELECT id, grade_details FROM excel_submission")
for r in cursor.fetchall():
    print(f"ID: {r[0]}")
    print(f"Grade Details: {repr(r[1][:200])}")
    print("="*40)
conn.close()
