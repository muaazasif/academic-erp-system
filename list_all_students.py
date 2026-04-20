from app import app, db, Student
with app.app_context():
    print(f"Current DB URI: {app.config['SQLALCHEMY_DATABASE_URI']}")
    students = Student.query.all()
    print(f"Total students in DB: {len(students)}")
    for s in students:
        print(f"ID: {s.student_id}, Name: {s.name}")
