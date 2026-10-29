from app import app, db
from sqlalchemy import text

with app.app_context():
    print("=== READ-ONLY BACKUP INVESTIGATION REPORT ===")
    
    # 1. Confirm table exists and row count
    try:
        res = db.session.execute(text("SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_schema = 'public' AND table_name = 'attendance_backup_20261028');"))
        exists = res.scalar()
        print(f"1. Table 'attendance_backup_20261028' exists: {exists}")
        
        if exists:
            count_res = db.session.execute(text("SELECT count(*) FROM public.attendance_backup_20261028;"))
            row_count = count_res.scalar()
            print(f"   Row count: {row_count}")
        else:
            print("   Backup table does not exist.")
            exit(0)
    except Exception as e:
        print(f"Error checking table existence/count: {e}")
        exit(1)

    # 2. Inspect column names and data types for backup table
    print("\n2. Columns and Data Types in 'attendance_backup_20261028':")
    backup_cols = {}
    try:
        col_res = db.session.execute(text("""
            SELECT column_name, data_type 
            FROM information_schema.columns 
            WHERE table_schema = 'public' AND table_name = 'attendance_backup_20261028'
            ORDER BY ordinal_position;
        """))
        for row in col_res:
            backup_cols[row[0]] = row[1]
            print(f"   - {row[0]}: {row[1]}")
    except Exception as e:
        print(f"Error getting columns: {e}")

    # 3. Compare columns with current 'attendance' table
    print("\n3. Comparison with current 'attendance' table columns:")
    current_cols = {}
    try:
        curr_res = db.session.execute(text("""
            SELECT column_name, data_type 
            FROM information_schema.columns 
            WHERE table_schema = 'public' AND table_name = 'attendance'
            ORDER BY ordinal_position;
        """))
        for row in curr_res:
            current_cols[row[0]] = row[1]
            print(f"   - {row[0]}: {row[1]}")
            
        print("\n   Column comparison:")
        only_in_backup = set(backup_cols.keys()) - set(current_cols.keys())
        only_in_current = set(current_cols.keys()) - set(backup_cols.keys())
        common = set(backup_cols.keys()).intersection(set(current_cols.keys()))
        print(f"   - Common columns: {common}")
        print(f"   - Only in backup: {only_in_backup if only_in_backup else 'None'}")
        print(f"   - Only in current: {only_in_current if only_in_current else 'None'}")
    except Exception as e:
        print(f"Error comparing columns: {e}")

    # 4. Report attendance records and date range
    print("\n4. Attendance Date Range:")
    try:
        # Check if 'date' column exists
        if 'date' in backup_cols:
            date_res = db.session.execute(text("SELECT min(date), max(date) FROM public.attendance_backup_20261028;"))
            min_d, max_d = date_res.fetchone()
            print(f"   - Minimum date: {min_d}")
            print(f"   - Maximum date: {max_d}")
        else:
            print("   - 'date' column not found in backup table.")
    except Exception as e:
        print(f"Error getting date range: {e}")

    # 5. Check for duplicate rows or constraints preventing safe restore
    print("\n5. Duplicate rows / Integrity check:")
    try:
        dup_res = db.session.execute(text("""
            SELECT student_id, date, count(*) 
            FROM public.attendance_backup_20261028 
            GROUP BY student_id, date 
            HAVING count(*) > 1;
        """))
        dups = dup_res.fetchall()
        print(f"   - Duplicate (student_id, date) pairs found: {len(dups)}")
    except Exception as e:
        print(f"Error checking duplicates: {e}")

    print("=== END OF INVESTIGATION REPORT ===")
