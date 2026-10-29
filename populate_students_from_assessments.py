from app import app, db, Assessment, Student
from werkzeug.security import generate_password_hash

with app.app_context():
    assessments = Assessment.query.all()
    print(f"Total assessments found: {len(assessments)}")
    existing_ids = {s.student_id for s in Student.query.all()}
    count = 0
    for a in assessments:
        if not a.student_id:
            continue
        sid = str(a.student_id).strip()
        if sid not in existing_ids:
            s = Student(
                student_id=sid,
                name=a.name or 'Student',
                batch=a.batch or 'Bano Qabil 5.0 Fall',
                campus=a.campus or '',
                timings=a.timings or '',
                gender=a.gender or 'Female',
                password_hash=generate_password_hash('student123')
            )
            db.session.add(s)
            existing_ids.add(sid)
            count += 1
            if count % 50 == 0:
                db.session.commit()
                print(f"Committed {count} students...")
    db.session.commit()
    print(f"✅ Successfully created {count} student accounts from assessments!")
    print(f"Total students now in database: {Student.query.count()}")
