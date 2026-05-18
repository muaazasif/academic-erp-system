import sqlite3
import pandas as pd
import json
import re
from io import BytesIO

def get_sample_data_sql():
    """Returns the SQL to initialize the sample database"""
    return """
    CREATE TABLE Students (
        id INTEGER PRIMARY KEY,
        name TEXT,
        city TEXT,
        joining_date DATE
    );
    
    CREATE TABLE Courses (
        id INTEGER PRIMARY KEY,
        title TEXT,
        duration_months INTEGER,
        fee INTEGER
    );
    
    CREATE TABLE Enrollments (
        student_id INTEGER,
        course_id INTEGER,
        enrollment_date DATE,
        FOREIGN KEY(student_id) REFERENCES Students(id),
        FOREIGN KEY(course_id) REFERENCES Courses(id)
    );
    
    INSERT INTO Students (id, name, city, joining_date) VALUES 
    (1, 'Ahmed Khan', 'Karachi', '2024-01-15'),
    (2, 'Sara Ahmed', 'Lahore', '2024-02-20'),
    (3, 'Bilal Sheikh', 'Karachi', '2024-03-10'),
    (4, 'Ayesha Malik', 'Islamabad', '2024-05-05'),
    (5, 'Omar Raza', 'Lahore', '2024-05-15');
    
    INSERT INTO Courses (id, title, duration_months, fee) VALUES 
    (101, 'Python Basics', 3, 5000),
    (102, 'Excel Mastery', 2, 3000),
    (103, 'SQL for Data', 4, 7000),
    (104, 'Web Dev', 6, 12000);
    
    INSERT INTO Enrollments (student_id, course_id, enrollment_date) VALUES 
    (1, 101, '2024-01-16'),
    (2, 102, '2024-02-21'),
    (3, 101, '2024-03-11'),
    (4, 103, '2024-05-06'),
    (5, 104, '2024-05-16');
    """

def get_sql_assignment_questions():
    """Returns the 10 tasks for the SQL Basic Practical assignment"""
    return [
        {
            "id": 1,
            "task": "Select all columns from the Students table.",
            "expected_query": "SELECT * FROM Students"
        },
        {
            "id": 2,
            "task": "Select the first 2 courses from the Courses table (use LIMIT).",
            "expected_query": "SELECT * FROM Courses LIMIT 2"
        },
        {
            "id": 3,
            "task": "Find all students who live in 'Karachi'.",
            "expected_query": "SELECT * FROM Students WHERE city = 'Karachi'"
        },
        {
            "id": 4,
            "task": "Find all students whose name starts with 'A'.",
            "expected_query": "SELECT * FROM Students WHERE name LIKE 'A%'"
        },
        {
            "id": 5,
            "task": "Find courses with fee greater than 5000 AND duration less than 12 months.",
            "expected_query": "SELECT * FROM Courses WHERE fee > 5000 AND duration_months < 12"
        },
        {
            "id": 6,
            "task": "Find students who live in either 'Lahore' OR 'Islamabad'.",
            "expected_query": "SELECT * FROM Students WHERE city = 'Lahore' OR city = 'Islamabad'"
        },
        {
            "id": 7,
            "task": "Select all courses and sort them by fee in descending order (highest first).",
            "expected_query": "SELECT * FROM Courses ORDER BY fee DESC"
        },
        {
            "id": 8,
            "task": "Count how many students are in each city. (Hint: GROUP BY city)",
            "expected_query": "SELECT city, COUNT(*) FROM Students GROUP BY city"
        },
        {
            "id": 9,
            "task": "Find students who joined in the month of May (05). Use strftime('%m', joining_date).",
            "expected_query": "SELECT * FROM Students WHERE strftime('%m', joining_date) = '05'"
        },
        {
            "id": 10,
            "task": "Show student names and the titles of the courses they are enrolled in. (Hint: INNER JOIN Students and Courses using Enrollments table)",
            "expected_query": "SELECT Students.name, Courses.title FROM Students INNER JOIN Enrollments ON Students.id = Enrollments.student_id INNER JOIN Courses ON Enrollments.course_id = Courses.id"
        }
    ]

def get_sample_data_as_excel(sample_sql=None):
    """Returns an Excel file as BytesIO containing all sample tables"""
    if not sample_sql:
        sample_sql = get_sample_data_sql()
        
    conn = sqlite3.connect(':memory:')
    cursor = conn.cursor()
    
    try:
        # Initialize schema and data
        cursor.executescript(sample_sql)
        conn.commit()
        
        # Get list of all tables in the database
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [row[0] for row in cursor.fetchall() if row[0] != 'sqlite_sequence']
        
        # Create Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for table_name in tables:
                df = pd.read_sql_query(f"SELECT * FROM {table_name}", conn)
                df.to_excel(writer, index=False, sheet_name=table_name)
        
        output.seek(0)
        return output
    finally:
        conn.close()

def grade_sql_submission(student_queries, questions, sample_sql=None):
    """
    Grades a student's SQL submission.
    student_queries: list of strings (queries for each task)
    questions: list of dicts with 'id', 'task', 'expected_query'
    sample_sql: SQL to initialize the database
    """
    if not sample_sql:
        sample_sql = get_sample_data_sql()
        
    total_score = 0
    max_score = len(questions)
    details = []
    
    # Student's database (persists across tasks in case of VIEW creation)
    conn_std = sqlite3.connect(':memory:')
    conn_std.executescript(sample_sql)
    
    # Reference database (for baseline results)
    conn_ref = sqlite3.connect(':memory:')
    conn_ref.executescript(sample_sql)
    
    try:
        for i, q in enumerate(questions):
            task_id = q['id']
            expected_sql = q['expected_query']
            student_sql = student_queries[i] if i < len(student_queries) else ""
            
            task_result = {
                "id": task_id,
                "task": q['task'],
                "correct": False,
                "error": None
            }
            
            if not student_sql or student_sql.strip() == "":
                task_result["error"] = "No query provided."
                details.append(task_result)
                continue
            
            try:
                # 1. Handle non-SELECT queries (CREATE VIEW, etc.)
                is_create_view = bool(re.search(r"CREATE\s+VIEW", student_sql, re.IGNORECASE))
                
                if is_create_view:
                    # Execute student's CREATE VIEW
                    conn_std.execute(student_sql)
                    conn_std.commit()
                    
                    # Extract view name
                    view_match = re.search(r"CREATE\s+VIEW\s+(\w+)", student_sql, re.IGNORECASE)
                    if view_match:
                        view_name = view_match.group(1)
                        student_df = pd.read_sql_query(f"SELECT * FROM {view_name}", conn_std)
                    else:
                        student_df = pd.DataFrame() # Fallback
                else:
                    # Regular SELECT query
                    student_df = pd.read_sql_query(student_sql, conn_std)
                
                # 2. Get expected result
                # If expected_sql is a CREATE VIEW, we execute it on ref DB first
                if bool(re.search(r"CREATE\s+VIEW", expected_sql, re.IGNORECASE)):
                    conn_ref.execute(expected_sql)
                    conn_ref.commit()
                    view_match_ref = re.search(r"CREATE\s+VIEW\s+(\w+)", expected_sql, re.IGNORECASE)
                    if view_match_ref:
                        view_name_ref = view_match_ref.group(1)
                        expected_df = pd.read_sql_query(f"SELECT * FROM {view_name_ref}", conn_ref)
                    else:
                        expected_df = pd.DataFrame()
                else:
                    expected_df = pd.read_sql_query(expected_sql, conn_ref)
                
                # 3. Compare
                if expected_df.equals(student_df):
                    task_result["correct"] = True
                    total_score += 1
                else:
                    # Normalize and compare
                    if not expected_df.empty and not student_df.empty:
                        # Check column count
                        if len(expected_df.columns) == len(student_df.columns):
                            # Sort for comparison if no ORDER BY
                            if "ORDER BY" not in expected_sql.upper():
                                # Try sorting by all columns
                                exp_sorted = expected_df.sort_values(by=list(expected_df.columns)).reset_index(drop=True)
                                std_sorted = student_df.sort_values(by=list(student_df.columns)).reset_index(drop=True)
                                if exp_sorted.equals(std_sorted):
                                    task_result["correct"] = True
                                    total_score += 1
                                else:
                                    task_result["error"] = "Results do not match expected output."
                            else:
                                task_result["error"] = "Results do not match (Ordering matters)."
                        else:
                            task_result["error"] = f"Column mismatch. Expected {len(expected_df.columns)}, got {len(student_df.columns)}."
                    else:
                        task_result["error"] = "Results do not match expected output."
            
            except Exception as e:
                task_result["error"] = f"SQL Error: {str(e)}"
            
            details.append(task_result)
            
    finally:
        conn_std.close()
        conn_ref.close()
        
    return {
        "score": total_score,
        "max": max_score,
        "percentage": (total_score / max_score) * 100 if max_score > 0 else 0,
        "details": details
    }
