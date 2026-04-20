from app import app, db, Student, Admin
import pandas as pd

with app.app_context():
    print("--- Admin Users ---")
    admins = Admin.query.all()
    for a in admins:
        print(f"Username: {a.username}")
    
    print("\n--- Student Users (Last 20) ---")
    students = Student.query.order_by(Student.created_at.desc()).limit(20).all()
    for s in students:
        print(f"ID: {s.student_id}, Name: {s.name}, Created At: {s.created_at}")
    
    # Check for duplicates
    from sqlalchemy import func
    duplicates = db.session.query(Student.student_id, func.count(Student.student_id)).group_by(Student.student_id).having(func.count(Student.student_id) > 1).all()
    if duplicates:
        print("\n--- Duplicate Student IDs found! ---")
        for d in duplicates:
            print(f"ID: {d[0]}, Count: {d[1]}")
    else:
        print("\nNo duplicate Student IDs found.")
