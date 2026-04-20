import sqlite3

db_path = 'instance/bq-erp.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

print(f"--- Tables in {db_path} ---")
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = cursor.fetchall()
for table in tables:
    print(f"Table: {table[0]}")
    cursor.execute(f"SELECT COUNT(*) FROM {table[0]}")
    count = cursor.fetchone()[0]
    print(f"  Count: {count}")
    
    if count > 0:
        cursor.execute(f"SELECT * FROM {table[0]} LIMIT 5")
        rows = cursor.fetchall()
        for row in rows:
            print(f"    {row}")

conn.close()
