import sqlite3

conn = sqlite3.connect('bq-erp.db')
cursor = conn.cursor()

print("--- TABLES ---")
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
for row in cursor.fetchall():
    print(row[0])

print("\n--- STUDENTS MATCHING 'muz' / 'maaz' / 'moiz' ---")
cursor.execute("SELECT id, name FROM students WHERE name LIKE '%muz%' OR name LIKE '%maaz%' OR name LIKE '%moiz%' OR name LIKE '%muzay%'")
students = cursor.fetchall()
for s in students:
    print(s)

print("\n--- EXCEL ASSIGNMENTS TABLES / DATA ---")
# Check if there are excel submission tables
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%excel%'")
for row in cursor.fetchall():
    print("Excel table:", row[0])
    cursor.execute(f"PRAGMA table_info({row[0]})")
    print("  Columns:", [col[1] for col in cursor.fetchall()])

conn.close()
