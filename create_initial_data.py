import os
import json
from app import app, db, Admin, Student, ExcelSkillsAssignment, SQLSkillsAssignment
from werkzeug.security import generate_password_hash
from datetime import datetime, timedelta
from sql_grader import get_sql_assignment_questions

def create_initial_data():
    # Ensure the instance directory exists
    os.makedirs(app.instance_path, exist_ok=True)

    with app.app_context():
        # Ensure all tables exist
        db.create_all()
        
        # 1. Check if admin user already exists
        admin_exists = Admin.query.filter_by(username='admin').first()
        # ... rest of admin creation ...

        if not admin_exists:
            # Create default admin user
            admin = Admin(
                username='admin',
                password_hash=generate_password_hash('admin123')
            )
            db.session.add(admin)
            db.session.commit()
            print("✅ Admin user created successfully!")
        else:
            print("✅ Admin user already exists.")

        # 2. Optionally create a sample student
        student_exists = Student.query.filter_by(student_id='101').first()

        if not student_exists:
            student = Student(
                student_id='101',
                name='John Doe',
                password_hash=generate_password_hash('student123')
            )
            db.session.add(student)
            db.session.commit()
            print("✅ Sample student created successfully!")
        else:
            print("✅ Sample student already exists.")

        # 3. Create Excel Skill assignments
        skills = [
            {
                "title": "Excel Skill 1: Formulas & Basics",
                "description": "Master the basics: 1. VLOOKUP (2 marks). 2. SUMIF & COUNTIF (2 marks). 3. Text Functions (LEFT, RIGHT, MID) (2 marks). 4. Nested IF (2 marks). 5. Complex Challenge (2 marks). Total 10 marks. AI will grade and provide instant feedback."
            },
            {
                "title": "Excel Skill 2: Data Validation & Named Manager",
                "description": "Master Workbook management: 1. Create Named Ranges (Name Manager). 2. Basic Dropdowns. 3. Advanced Dependent Dropdowns. 4. Data Validation (Numbers, Dates, Text). Total 10 marks. AI will provide instant feedback on mistakes."
            },
            {
                "title": "Excel Skill 3: Data Cleaning & Power Query",
                "description": "Master Data Preparation: 1. Cleaning Raw Data (Spaces, Case, Duplicates). 2. Text-to-Columns & Flash Fill. 3. Power Query Basics (Transforming & Loading). Total 10 marks. AI will check for clean data and proper transformations."
            },
            {
                "title": "Excel Skill 4: Advanced LOOKUP & Aggregation",
                "description": "Master Data Relationships: 1. LOOKUP Function (Vector/Array) (2.5 marks). 2. Advanced SUMIFS (2.5 marks). 3. COUNTIFS & Relationships (2.5 marks). 4. Integrated Challenge (2.5 marks). Total 10 marks. (Note: specifically NOT VLOOKUP/XLOOKUP)"
            }
        ]

        for skill_data in skills:
            skill = ExcelSkillsAssignment.query.filter_by(title=skill_data["title"]).first()
            if not skill:
                new_skill = ExcelSkillsAssignment(
                    title=skill_data["title"],
                    description=skill_data["description"],
                    created_at=datetime.now(),
                    deadline=datetime.now() + timedelta(days=14),
                    is_active=True
                )
                db.session.add(new_skill)
                print(f"✅ {skill_data['title']} created!")
            else:
                if not skill.is_active:
                    skill.is_active = True
                print(f"✅ {skill_data['title']} exists/activated.")
        
        # 4. Create SQL Skills Assignments
        sql_assignments_data = [
            {
                "title": "SQL Basic Practical",
                "description": "Master basic SQL queries: SELECT, LIMIT, WHERE, LIKE, GROUP BY, ORDER BY, and INNER JOIN. Total 10 marks. AI will grade your queries instantly.",
                "questions": get_sql_assignment_questions()
            },
            {
                "title": "SQL Medium Level: Views & Joins",
                "description": "Master Medium level SQL: CREATE VIEW, multi-table INNER JOIN, LEFT JOIN, and Subqueries. Total 10 marks.",
                "questions": [
                    {"id": 1, "task": "Create a VIEW named 'StudentEnrollments' that shows Student names and their Course titles.", "expected_query": "SELECT Students.name, Courses.title FROM Students INNER JOIN Enrollments ON Students.id = Enrollments.student_id INNER JOIN Courses ON Enrollments.course_id = Courses.id"},
                    {"id": 2, "task": "Select all columns from the 'StudentEnrollments' view.", "expected_query": "SELECT * FROM StudentEnrollments"},
                    {"id": 3, "task": "List all students and their enrollment date (if any) using a LEFT JOIN.", "expected_query": "SELECT Students.name, Enrollments.enrollment_date FROM Students LEFT JOIN Enrollments ON Students.id = Enrollments.student_id"},
                    {"id": 4, "task": "Find the total fee collected from all enrollments. (SUM of course fees)", "expected_query": "SELECT SUM(Courses.fee) FROM Enrollments INNER JOIN Courses ON Enrollments.course_id = Courses.id"},
                    {"id": 5, "task": "Find cities where more than 1 student resides. (Use HAVING)", "expected_query": "SELECT city FROM Students GROUP BY city HAVING COUNT(*) > 1"},
                    {"id": 6, "task": "Find students who joined after 'Ahmed Khan' (id=1).", "expected_query": "SELECT * FROM Students WHERE joining_date > (SELECT joining_date FROM Students WHERE id = 1)"},
                    {"id": 7, "task": "Show course titles and the number of students enrolled in each.", "expected_query": "SELECT Courses.title, COUNT(Enrollments.student_id) FROM Courses LEFT JOIN Enrollments ON Courses.id = Enrollments.course_id GROUP BY Courses.title"},
                    {"id": 8, "task": "Find the most expensive course title and its fee.", "expected_query": "SELECT title, fee FROM Courses ORDER BY fee DESC LIMIT 1"},
                    {"id": 9, "task": "Get the names of students enrolled in 'Python Basics'.", "expected_query": "SELECT name FROM Students WHERE id IN (SELECT student_id FROM Enrollments WHERE course_id = 101)"},
                    {"id": 10, "task": "Create a view named 'KarachiStudents' for students living in Karachi.", "expected_query": "SELECT * FROM Students WHERE city = 'Karachi'"}
                ]
            },
            {
                "title": "SQL Advanced Level: Complex Queries",
                "description": "Master Advanced SQL: Complex CTEs, Subqueries in FROM clause, and CASE statements. Total 10 marks.",
                "questions": [
                    {"id": 1, "task": "Create a VIEW 'DetailedReport' joining Students, Enrollments, and Courses with all details.", "expected_query": "SELECT Students.name, Students.city, Courses.title, Courses.fee, Enrollments.enrollment_date FROM Students JOIN Enrollments ON Students.id = Enrollments.student_id JOIN Courses ON Enrollments.course_id = Courses.id"},
                    {"id": 2, "task": "Find students who are enrolled in more than 1 course.", "expected_query": "SELECT name FROM Students WHERE id IN (SELECT student_id FROM Enrollments GROUP BY student_id HAVING COUNT(*) > 1)"},
                    {"id": 3, "task": "Use a CTE (WITH clause) to list students from 'Karachi' and their total course fees.", "expected_query": "WITH StudentFees AS (SELECT student_id, SUM(fee) as total FROM Enrollments JOIN Courses ON Enrollments.course_id = Courses.id GROUP BY student_id) SELECT name, total FROM Students JOIN StudentFees ON Students.id = StudentFees.student_id WHERE city = 'Karachi'"},
                    {"id": 4, "task": "Find the top 2 highest paying students and their names.", "expected_query": "SELECT name, SUM(fee) FROM Students JOIN Enrollments ON Students.id = Enrollments.student_id JOIN Courses ON Enrollments.course_id = Courses.id GROUP BY name ORDER BY SUM(fee) DESC LIMIT 2"},
                    {"id": 5, "task": "Find courses that have no enrollments.", "expected_query": "SELECT title FROM Courses WHERE id NOT IN (SELECT course_id FROM Enrollments)"},
                    {"id": 6, "task": "Show student names and a column 'Status' which is 'Karachi Resident' if they live in Karachi, otherwise 'Other'.", "expected_query": "SELECT name, CASE WHEN city = 'Karachi' THEN 'Karachi Resident' ELSE 'Other' END as Status FROM Students"},
                    {"id": 7, "task": "Find the student who enrolled first in any course.", "expected_query": "SELECT name FROM Students JOIN Enrollments ON Students.id = Enrollments.student_id ORDER BY enrollment_date ASC LIMIT 1"},
                    {"id": 8, "task": "Get a list of all cities and the total revenue from each city.", "expected_query": "SELECT Students.city, SUM(Courses.fee) FROM Students JOIN Enrollments ON Students.id = Enrollments.student_id JOIN Courses ON Enrollments.course_id = Courses.id GROUP BY Students.city"},
                    {"id": 9, "task": "Find students who live in the same city as 'Sara Ahmed'.", "expected_query": "SELECT name FROM Students WHERE city = (SELECT city FROM Students WHERE name = 'Sara Ahmed') AND name != 'Sara Ahmed'"},
                    {"id": 10, "task": "Calculate the percentage of total students that live in each city.", "expected_query": "SELECT city, COUNT(*)*100.0 / (SELECT COUNT(*) FROM Students) FROM Students GROUP BY city"}
                ]
            }
        ]

        for sql_data in sql_assignments_data:
            sql_assignment = SQLSkillsAssignment.query.filter_by(title=sql_data["title"]).first()
            if not sql_assignment:
                new_sql = SQLSkillsAssignment(
                    title=sql_data["title"],
                    description=sql_data["description"],
                    questions_json=json.dumps(sql_data["questions"]),
                    created_at=datetime.now(),
                    deadline=datetime.now() + timedelta(days=14),
                    is_active=True
                )
                db.session.add(new_sql)
                print(f"✅ {sql_data['title']} created!")
            else:
                if not sql_assignment.is_active:
                    sql_assignment.is_active = True
                print(f"✅ {sql_data['title']} exists/activated.")

        db.session.commit()

if __name__ == '__main__':
    create_initial_data()
