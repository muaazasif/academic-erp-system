import sqlite3
import os

for db_path in ['erp_system.db', 'bq-erp.db', 'instance/erp_system.db', 'instance/bq-erp.db']:
    if os.path.exists(db_path):
        print(f"Checking {db_path}...")
        try:
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()
            print("Tables:", tables)
            for t in tables:
                t_name = t[0]
                cursor.execute(f"SELECT COUNT(*) FROM {t_name}")
                print(f"  Table {t_name}: {cursor.fetchone()[0]} rows")
            conn.close()
        except Exception as e:
            print(f"Error: {e}")
