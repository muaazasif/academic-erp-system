import os
from app import app, db, Admin, Student, ExcelSkillsAssignment
from werkzeug.security import generate_password_hash
from datetime import datetime, timedelta

def create_initial_data():
    # Ensure the instance directory exists
    os.makedirs(app.instance_path, exist_ok=True)

    with app.app_context():
        # 1. Check if admin user already exists
        admin_exists = Admin.query.filter_by(username='admin').first()

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
        
        db.session.commit()

if __name__ == '__main__':
    create_initial_data()