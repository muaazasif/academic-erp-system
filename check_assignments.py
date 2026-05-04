
from app import app, db, ExcelSkillsAssignment
from datetime import datetime, timedelta

with app.app_context():
    print(f"Checking database: {app.config['SQLALCHEMY_DATABASE_URI']}")
    
    # 1. Cleanup old assignments to avoid confusion (Optional, but let's ensure names match)
    # Existing names we might have used:
    old_titles = [
        "Advanced Excel: Data Validation & AI Grading",
        "Excel Skills: Data Validation & Named Manager"
    ]
    for old_t in old_titles:
        old_a = ExcelSkillsAssignment.query.filter_by(title=old_t).first()
        if old_a:
            db.session.delete(old_a)
            print(f"🗑️ Deleted old assignment: {old_t}")
    db.session.commit()

    # 2. Force creation of 1 and 2
    for title, desc in [
        ("Excel Skill 1: Formulas & Basics", "Master the basics: 1. VLOOKUP (2 marks). 2. SUMIF & COUNTIF (2 marks). 3. Text Functions (LEFT, RIGHT, MID) (2 marks). 4. Nested IF (2 marks). 5. Complex Challenge (2 marks). Total 10 marks. AI will grade and provide instant feedback."),
        ("Excel Skill 2: Data Validation & Named Manager", "Master Workbook management: 1. Create Named Ranges (Name Manager). 2. Basic Dropdowns. 3. Advanced Dependent Dropdowns. 4. Data Validation (Numbers, Dates, Text). Total 10 marks. AI will provide instant feedback on mistakes.")
    ]:
        exists = ExcelSkillsAssignment.query.filter_by(title=title).first()
        if not exists:
            new_a = ExcelSkillsAssignment(
                title=title,
                description=desc,
                created_at=datetime.now(),
                deadline=datetime.now() + timedelta(days=14),
                is_active=True
            )
            db.session.add(new_a)
            db.session.commit()
            print(f"✅ Created: {title}")
        else:
            print(f"✅ Already exists: {title}")

    assignments = ExcelSkillsAssignment.query.all()
    print(f"\nFinal Total Assignments in {app.config['SQLALCHEMY_DATABASE_URI']}: {len(assignments)}")
    for a in assignments:
        print(f"ID: {a.id} | Title: {a.title}")
