import re
from app import app, db, Student
from werkzeug.security import generate_password_hash
from sqlalchemy import text

with app.app_context():
    print("Reading attendance_import_with_locations.sql to extract student IDs...")
    with open("attendance_import_with_locations.sql", "r", encoding="utf-8") as f:
        content = f.read()

    # Extract all student_ids from VALUES tuples
    # Values pattern: ('2026-04-12', '504625', 'Marium Gul', ...)
    pattern = r"\('\d{4}-\d{2}-\d{2}',\s*'([^']+)'"
    student_ids = set(re.findall(pattern, content))
    print(f"Found {len(student_ids)} unique student IDs in attendance SQL.")

    existing_students = {s.student_id for s in Student.query.all()}
    missing_ids = student_ids - existing_students
    print(f"Missing student IDs in public.student: {len(missing_ids)}")
    print(list(missing_ids)[:30])

    if missing_ids:
        print("Inserting missing stub student records...")
        for sid in missing_ids:
            s = Student(
                student_id=sid,
                name=f"Student {sid}",
                batch="Bano Qabil 5.0 Fall",
                gender="Female",
                password_hash=generate_password_hash("student123")
            )
            db.session.add(s)
        db.session.commit()
        print("✅ Missing student records inserted successfully!")

    # Now execute the attendance import SQL script
    try:
        print("Executing attendance import script...")
        db.session.execute(text(content))
        db.session.commit()
        print("✅ Attendance import successfully executed and committed!")
        
        count_res = db.session.execute(text("SELECT count(*) FROM public.attendance;"))
        print(f"Total attendance records now in database: {count_res.scalar()}")
    except Exception as e:
        db.session.rollback()
        print(f"❌ Attendance import failed: {e}")
