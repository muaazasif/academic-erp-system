import sqlite3

db_path = 'instance/erp_system.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

cursor.execute("PRAGMA table_info(student);")
columns = cursor.fetchall()
for col in columns:
    print(col)

conn.close()
