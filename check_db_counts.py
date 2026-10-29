from app import app, db, Student, Assessment, Attendance, ExcelSubmission, SQLSubmission

with app.app_context():
    try:
        print("--- READ-ONLY DATABASE VERIFICATION REPORT ---")
        print(f"Students count: {Student.query.count()}")
        print(f"Assessments count: {Assessment.query.count()}")
        print(f"Attendance count: {Attendance.query.count()}")
        print(f"Excel Submissions count: {ExcelSubmission.query.count()}")
        print(f"SQL Submissions count: {SQLSubmission.query.count()}")
        print("----------------------------------------------")
    except Exception as e:
        print(f"Error querying database: {e}")
