from app import app, db, Student, Assessment, Batch

with app.app_context():
    students = Student.query.all()
    updated = 0
    for s in students:
        assessment = Assessment.query.filter_by(student_id=str(s.student_id)).first()
        if not assessment:
            assessment = Assessment.query.filter_by(email_address=s.student_id).first()
        if assessment:
            s.batch = assessment.batch
            s.campus = assessment.campus
            s.timings = assessment.timings
            updated += 1
    db.session.commit()
    print(f"✅ Linked {updated} students to their respective batches from assessments!")
