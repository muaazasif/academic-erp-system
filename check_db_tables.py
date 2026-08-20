import sqlite3
import os

def check_db():
    dbs = ['instance/bq-erp.db', 'instance/erp_system.db']
    for db_path in dbs:
        if os.path.exists(db_path):
            print(f"\n--- Checking {db_path} ---")
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            tables = [t[0] for t in cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")]
            print(f"Tables: {tables}")
            
            for table in tables:
                if 'mid' in table.lower() or 'assign' in table.lower():
                    try:
                        count = cursor.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                        print(f"Table '{table}' has {count} rows")
                        if count > 0:
                            # Print column names
                            cols = [c[1] for c in cursor.execute(f"PRAGMA table_info({table})")]
                            print(f"Columns: {cols}")
                            rows = cursor.execute(f"SELECT * FROM {table} LIMIT 1").fetchall()
                            print(f"Sample row from '{table}': {rows[0]}")
                    except:
                        pass
            conn.close()

if __name__ == "__main__":
    check_db()
