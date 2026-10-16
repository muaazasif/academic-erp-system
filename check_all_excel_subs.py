from app import app, db, Student, ExcelSubmission, ExcelSkillsAssignment
with app.app_context():
    print("--- ALL EXCEL SUBMISSIONS ---")
    subs = ExcelSubmission.query.all()
    print(f"Total ExcelSubmissions: {len(subs)}")
    for sub in subs:
        student = Student.query.filter_by(student_id=sub.student_id).first()
        name = student.name if student else "Unknown"
        print(f"Student ID: {sub.student_id}, Name: {name}, Skill ID: {sub.assignment_id}, Score: {sub.score}, Total: {sub.total_marks}, Grade: {sub.grade}")

    print("\n--- ALL STUDENTS WITH 'muz' or similar ---")
    students = Student.query.filter(
        (Student.name.ilike('%muz%')) | 
        (Student.name.ilike('%muza%')) |
        (Student.name.ilike('%maaz%')) | 
        (Student.name.ilike('%moiz%')) |
        (Student.name.ilike('%muzaffar%')) |
        (Student.name.ilike('%muzay%'))
    ).all()
    for s in students:
        print(f"ID: {s.student_id}, Name: {s.name}")
        s_subs = ExcelSubmission.query.filter_by(student_id=s.student_id).all()
        for sub in s_subs:
            print(f"  Skill {sub.assignment_id}: Score={sub.score}/{sub.total_marks}, Grade={sub.grade}")
