from app import app, db, Student, ExcelSubmission, ExcelSkillsAssignment, ExcelAssignment
with app.app_context():
    # Find students matching muz or maaz or muzay
    query_str = "%muz%"
    students = Student.query.filter(
        (Student.name.ilike('%muz%')) | 
        (Student.name.ilike('%maaz%')) | 
        (Student.name.ilike('%moiz%')) |
        (Student.name.ilike('%muzay%'))
    ).all()
    
    print(f"Found {len(students)} matching students:")
    for s in students:
        print(f"ID: {s.student_id}, Name: {s.name}")
        
        # Check ExcelSubmission for this student
        submissions = ExcelSubmission.query.filter_by(student_id=s.student_id).all()
        print(f"  ExcelSubmissions: {len(submissions)}")
        for sub in submissions:
            print(f"    Assignment ID/Skill: {sub.assignment_id}, Score/Marks: {sub.score}, Total: {sub.total_marks}, Grade: {sub.grade}, Feedback: {sub.feedback}")
            
    # Also let's check ExcelAssignments (Skill 1, 5, 6)
    print("\n--- Excel Skills / Assignments ---")
    skills = ExcelSkillsAssignment.query.all()
    for sk in skills:
        print(sk.__dict__)
        
    excel_assignments = ExcelAssignment.query.all()
    for ea in excel_assignments:
        print(ea.__dict__)
