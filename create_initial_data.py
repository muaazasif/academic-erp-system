import os
import json
from app import app, db, Admin, Student, ExcelSkillsAssignment, SQLSkillsAssignment, MidTerm, BrainLabChallenge
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
                "description": "Master Excel Formulas: VLOOKUP, SUMIF, COUNTIF, IF. Students must implement these correctly to solve the business case study provided. Total 5 marks. AI will grade based on formula accuracy and output."
            },
            {
                "title": "Excel Skill 2: Data Validation & Named Manager",
                "description": "Master Workbook management: 1. Create Named Ranges (Name Manager). 2. Basic Dropdowns. 3. Advanced Dependent Dropdowns. 4. Data Validation (Numbers, Dates, Text). Total 5 marks. AI will provide instant feedback on mistakes."
            },
            {
                "title": "Excel Skill 3: Data Cleaning & Power Query",
                "description": "Master Data Preparation: 1. Cleaning Raw Data (Spaces, Case, Duplicates). 2. Text-to-Columns & Flash Fill. 3. Power Query Basics (Transforming & Loading). Total 5 marks. AI will check for clean data and proper transformations."
            },
            {
                "title": "Excel Skill 4: Advanced LOOKUP & Aggregation",
                "description": "Master Data Relationships: 1. LOOKUP Function (Vector/Array). 2. Advanced SUMIFS. 3. COUNTIFS & Relationships. 4. Integrated Challenge. Total 5 marks. (Note: specifically NOT VLOOKUP/XLOOKUP)"
            },
            {
                "title": "Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula",
                "description": "Complex Assignment: 1. VLOOKUP (Range/Exact), 2. SUMIF, 3. COUNTIF, 4. IF Formula. Total 5 marks. AI checks formula logic and data accuracy. Feedback provided on mistakes."
            },
            {
                "title": "Excel Skill 6: IFERROR, DATE, TEXT, AND & OR",
                "description": "Advanced Formulas & Logic: 1. IFERROR (1 mark). 2. DATE (1 mark). 3. TEXT (1 mark). 4. AND Condition (1 mark). 5. OR Condition (1 mark). Total 5 marks. AI checks formula logic and data accuracy."
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

        # 6. Create Randomized Midterm Exam
        midterm = MidTerm.query.filter_by(title="Randomized Midterm Exam").first()
        if not midterm:
            new_midterm = MidTerm(
                title="Randomized Midterm Exam",
                description="Comprehensive Exam covering Excel (Basic/Advanced), SQL, Power Query, and VBA. Each student receives 10 unique tasks from a pool of 100. AI Auto-graded.",
                total_sheets=100,
                sheets_per_student=10,
                created_by="admin",
                due_date=datetime.now() + timedelta(days=30)
            )
            db.session.add(new_midterm)
            print("✅ Randomized Midterm Exam created!")
        
        db.session.commit()

        # 7. Seed Brain Lab Challenges
        seed_brain_labs()

def seed_brain_labs():
    """Seed or update default Brain Lab challenges"""
    challenges = [
        {
            "challenge_number": 1,
            "title": "The Suspicious Salary",
            "scenario": "The manager says:\n‘The average salary of Karachi employees is extremely high.’\n\nBut something may be wrong.\n\nDON’T CALCULATE YET.\n\nWhat looks suspicious?",
            "dataset_json": json.dumps([
                {"Employee": "Ali", "City": "Karachi", "Salary": "45,000"},
                {"Employee": "Ahmed", "City": "Lahore", "Salary": "52,000"},
                {"Employee": "Sara", "City": "Karachi", "Salary": "48,000"},
                {"Employee": "Hamza", "City": "Karachi", "Salary": "450,000"},
                {"Employee": "Zain", "City": "Lahore", "Salary": "55,000"}
            ]),
            "think_prompt": "Examine the salary figures for Karachi employees. Which entry stands out as abnormal before doing any math?",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "Ali's salary is too low"},
                {"id": "b", "text": "Hamza's salary of 450,000 is an extreme outlier / data entry error"},
                {"id": "c", "text": "All Karachi salaries are incorrect"},
                {"id": "d", "text": "Zain is assigned to the wrong city"}
            ]),
            "verify_instructions": "Now unlock Excel. Use `=AVERAGE()` to calculate the average salary with Hamza included, then calculate it excluding Hamza. Notice the massive difference.",
            "final_question": "What caused the unusual average reported by the manager, and how does it prove why data observation matters?",
            "correct_reasoning": "Hamza's salary of 450,000 is an extreme outlier that disproportionately skews the arithmetic mean. In data analytics, you must always check for outliers before trusting summary statistics.",
            "xp": 100
        },
        {
            "challenge_number": 2,
            "title": "The Double Total",
            "scenario": "A regional sales report shows individual store sales, but the Grand Total at the bottom is displayed as 950,000.\n\nBefore opening Excel, what would you investigate first?",
            "dataset_json": json.dumps([
                {"Store": "North", "Manager": "Bilal", "Month": "Jan", "Sales": "120,000"},
                {"Store": "South", "Manager": "Ayesha", "Month": "Jan", "Sales": "95,000"},
                {"Store": "East", "Manager": "Fahad", "Month": "Jan", "Sales": "150,000"},
                {"Store": "West", "Manager": "Sana", "Month": "Jan", "Sales": "110,000"},
                {"Store": "Grand Total", "Manager": "All", "Month": "Jan", "Sales": "950,000"}
            ]),
            "think_prompt": "Look closely at the individual store sales (120k + 95k + 150k + 110k = 475k) and compare it with the displayed Grand Total of 950,000.",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "The Grand Total of 950,000 is exactly double the sum of the four stores (475,000), indicating double-counting"},
                {"id": "b", "text": "Sales are too high overall across all regions"},
                {"id": "c", "text": "Manager names are misspelled in the North region"},
                {"id": "d", "text": "The month column has formatting errors"}
            ]),
            "verify_instructions": "Now unlock Excel. Use `=SUM()` on individual store sales. Notice how the actual sum is 475,000, while the report states 950,000 (475,000 × 2).",
            "final_question": "Why did the grand total appear as 950,000 instead of 475,000?",
            "correct_reasoning": "The individual store sales correctly add up to 475,000. The reported Grand Total of 950,000 is exactly double (475,000 × 2), creating a double-counting mystery.",
            "xp": 100
        },
        {
            "challenge_number": 3,
            "title": "Find the Impossible Record",
            "scenario": "HR provided an employee roster for analysis.\n\nBefore calculating departmental averages or tenure, scan the data for logical impossibilities.",
            "dataset_json": json.dumps([
                {"ID": "101", "Name": "Usman", "Age": "28", "Salary": "60,000", "Joining Date": "2021-05-12"},
                {"ID": "102", "Name": "Mariam", "Age": "150", "Salary": "75,000", "Joining Date": "2019-03-01"},
                {"ID": "103", "Name": "Bilal", "Age": "32", "Salary": "-20,000", "Joining Date": "2022-01-15"},
                {"ID": "104", "Name": "Hira", "Age": "25", "Salary": "45,000", "Joining Date": "2028-09-10"}
            ]),
            "think_prompt": "Which employee records contain values that defy physical or logical reality?",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "Mariam (Age 150), Bilal (Negative salary), and Hira (Future joining date 2028)"},
                {"id": "b", "text": "All employees have completely realistic records"},
                {"id": "c", "text": "Usman joined too early in 2021"},
                {"id": "d", "text": "Only salary values are incorrect"}
            ]),
            "verify_instructions": "Now unlock Excel. Use conditional formatting (e.g., Highlight Cell Rules < 0 or > 100) to flag invalid numbers.",
            "final_question": "What types of data quality issues must be cleaned before performing descriptive statistics?",
            "correct_reasoning": "Data entry errors such as impossible ages (150), negative salaries (-20,000), and future joining dates (2028) severely distort calculations like AVERAGE and COUNT.",
            "xp": 120
        },
        {
            "challenge_number": 4,
            "title": "Who Is Actually #1?",
            "scenario": "Two sales representatives are claiming to be the top performer:\n\n* **Rep A = Kamran** — highest total sales volume\n* **Rep B = Nida** — highest average deal size\n\nThe sales manager asks:\n\n> **‘Who should we consider the top performer?’**\n\nBefore calculating anything, **what should leadership investigate first?**",
            "dataset_json": json.dumps([
                {"Rep": "Kamran (Rep A)", "Total Sales": "2,500,000", "Deals Closed": "50", "Avg Deal Size": "50,000"},
                {"Rep": "Nida (Rep B)", "Total Sales": "1,800,000", "Deals Closed": "15", "Avg Deal Size": "120,000"},
                {"Rep": "Zubair", "Total Sales": "2,100,000", "Deals Closed": "35", "Avg Deal Size": "60,000"}
            ]),
            "think_prompt": "Don't choose based only on the biggest number.\n\nThink about **what ‘top performer’ actually means** and which metric would be relevant to the business goal.",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "Rep A has higher total sales, while Rep B has a higher average deal size."},
                {"id": "b", "text": "Rep B has both the highest total sales and highest average deal size."},
                {"id": "c", "text": "Zubair has the highest performance on every metric."},
                {"id": "d", "text": "All three representatives have almost identical performance."}
            ]),
            "verify_instructions": "Now unlock Excel. Sort the table by Total Sales, then sort by Avg Deal Size. Observe how rankings change.",
            "final_question": "How should leadership define '#1 performer' in this context?",
            "correct_reasoning": "Performance metrics depend on business strategy. Kamran drives aggregate revenue, whereas Nida secures high-value clients with greater efficiency and less deal friction.",
            "xp": 120
        },
        {
            "challenge_number": 5,
            "title": "The Duplicate",
            "scenario": "A customer feedback dataset shows satisfaction scores. You notice customer names spelled slightly differently.\n\nHow does this affect unique customer counts?",
            "dataset_json": json.dumps([
                {"Customer": "Muhammad Ali", "City": "Lahore", "Score": "9"},
                {"Customer": "Mohammad Ali", "City": "Lahore", "Score": "8"},
                {"Customer": "Fatima Noor", "City": "Karachi", "Score": "10"},
                {"Customer": "Ahmed Raza", "City": "Islamabad", "Score": "7"}
            ]),
            "think_prompt": "What problem do 'Muhammad Ali' and 'Mohammad Ali' present when calculating unique customers?",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "They are duplicate records with minor spelling variations that inflate unique customer counts"},
                {"id": "b", "text": "They are completely independent customers"},
                {"id": "c", "text": "Their feedback scores will cancel each other out"},
                {"id": "d", "text": "No data discrepancy exists"}
            ]),
            "verify_instructions": "Now unlock Excel. Use Find & Replace or sorting to check for spelling consistency.",
            "final_question": "Why is text standardization crucial before performing unique entity counts?",
            "correct_reasoning": "Spelling variations cause automated systems to treat the same person as two distinct individuals, skewing customer retention and satisfaction metrics.",
            "xp": 110
        },
        {
            "challenge_number": 6,
            "title": "The Date Mystery",
            "scenario": "Monthly transaction reports are combined, but quarterly revenue formulas return #VALUE! errors.\n\nBefore running formulas, what do you notice about date formats?",
            "dataset_json": json.dumps([
                {"Transaction ID": "T101", "Date": "12/15/2025", "Amount": "15,000"},
                {"Transaction ID": "T102", "Date": "2025-12-16", "Amount": "22,000"},
                {"Transaction ID": "T103", "Date": "17-Dec-2025", "Amount": "18,000"},
                {"Transaction ID": "T104", "Date": "Jan 12 2026", "Amount": "30,000"}
            ]),
            "think_prompt": "Are all transaction dates formatted consistently across the rows?",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "Dates have mixed formats (MM/DD, YYYY-MM-DD, text month names), preventing Excel from recognizing them as serial numbers"},
                {"id": "b", "text": "All dates are in standard serial number format"},
                {"id": "c", "text": "The transaction amounts are negative"},
                {"id": "d", "text": "Transaction IDs are duplicated"}
            ]),
            "verify_instructions": "Now unlock Excel. Try sorting by date or applying `=MONTH()` to see which cells fail with `#VALUE!`.",
            "final_question": "Why did date aggregation and quarterly formulas fail?",
            "correct_reasoning": "Excel stores dates as sequential numbers. Mixed text and date formats prevent chronological sorting and date-based aggregation formulas.",
            "xp": 130
        },
        {
            "challenge_number": 7,
            "title": "The Salary Trap",
            "scenario": "A department of 5 employees has 4 salaries around 50,000 and 1 CEO salary at 1,000,000.\n\nWill the Mean (Average) or Median better represent the typical salary?",
            "dataset_json": json.dumps([
                {"Employee": "Staff 1", "Salary": "48,000"},
                {"Employee": "Staff 2", "Salary": "52,000"},
                {"Employee": "Staff 3", "Salary": "50,000"},
                {"Employee": "Staff 4", "Salary": "55,000"},
                {"Employee": "CEO", "Salary": "1,000,000"}
            ]),
            "think_prompt": "How will the CEO's salary impact the arithmetic mean versus the median?",
            "predict_options_json": json.dumps([
                {"id": "a", "text": "Mean will be pulled very high by the outlier, whereas Median will remain representative of the 4 staff"},
                {"id": "b", "text": "Median will be higher than Mean"},
                {"id": "c", "text": "Both statistical measures will be identical"},
                {"id": "d", "text": "Neither formula will work in Excel"}
            ]),
            "verify_instructions": "Now unlock Excel. Use `=AVERAGE()` and `=MEDIAN()` to compare both results.",
            "final_question": "When should an analyst choose Median over Average?",
            "correct_reasoning": "When data contains extreme outliers or skewed distributions, the median provides a true reflection of the typical central tendency without being distorted by extremes.",
            "xp": 130
        }
    ]

    for c_data in challenges:
        existing = BrainLabChallenge.query.filter_by(challenge_number=c_data["challenge_number"]).first()
        if existing:
            existing.title = c_data["title"]
            existing.scenario = c_data["scenario"]
            existing.dataset_json = c_data["dataset_json"]
            existing.think_prompt = c_data["think_prompt"]
            existing.predict_options_json = c_data["predict_options_json"]
            existing.verify_instructions = c_data["verify_instructions"]
            existing.final_question = c_data["final_question"]
            existing.correct_reasoning = c_data["correct_reasoning"]
            existing.xp = c_data["xp"]
        else:
            challenge = BrainLabChallenge(**c_data)
            db.session.add(challenge)
    db.session.commit()
    print("✅ Brain Lab challenges seeded & updated successfully!")

if __name__ == '__main__':
    create_initial_data()
