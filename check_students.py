from app import app, db, Student
with app.app_context():
    students = Student.query.all()
    print(f'Total students: {len(students)}')
    for s in students:
        print(f'ID:[{s.student_id}], Name:[{s.name}]')
