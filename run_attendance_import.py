from app import app, db
from sqlalchemy import text

with app.app_context():
    try:
        print("Executing attendance_import_with_locations.sql...")
        with open("attendance_import_with_locations.sql", "r", encoding="utf-8") as f:
            sql_script = f.read()
        
        # Execute script statements
        db.session.execute(text(sql_script))
        db.session.commit()
        print("✅ Attendance import successfully executed and committed!")
        
        # Verify row count
        count_res = db.session.execute(text("SELECT count(*) FROM public.attendance;"))
        print(f"Total attendance records now in database: {count_res.scalar()}")
    except Exception as e:
        db.session.rollback()
        print(f"❌ Attendance import failed and rolled back: {e}")
