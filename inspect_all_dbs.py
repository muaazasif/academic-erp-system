import os
from sqlalchemy import create_engine, MetaData, Table, select

db_files = [
    'instance/bq-erp.db',
    'instance/erp_system.db',
    'bq-erp.db',
    'erp_system.db'
]

for db_file in db_files:
    if os.path.exists(db_file):
        print(f"\n=== Inspecting {db_file} ===")
        try:
            engine = create_engine(f"sqlite:///{db_file}")
            conn = engine.connect()
            metadata = MetaData()
            metadata.reflect(bind=engine)
            
            if 'admin' in metadata.tables:
                admin_table = metadata.tables['admin']
                admins = conn.execute(select(admin_table)).fetchall()
                print(f"Admins: {len(admins)}")
                for a in admins:
                    # Depending on schema, columns might be id, username, password_hash
                    print(f"  - {a}")
            else:
                print("No 'admin' table found.")
                
            if 'student' in metadata.tables:
                student_table = metadata.tables['student']
                students_count = conn.execute(select(student_table)).rowcount # This doesn't work well for all dialects
                students = conn.execute(select(student_table).limit(5)).fetchall()
                # To get count properly
                from sqlalchemy import func
                count_query = select(func.count()).select_from(student_table)
                count = conn.execute(count_query).scalar()
                print(f"Students: {count}")
                for s in students:
                    print(f"  - {s}")
            else:
                print("No 'student' table found.")
                
            conn.close()
        except Exception as e:
            print(f"Error inspecting {db_file}: {e}")
    else:
        print(f"\nFile {db_file} does not exist.")
