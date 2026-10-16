from app import app, db, ExcelSkillsAssignment

with app.app_context():
    db.create_all()
    excel_skills = ExcelSkillsAssignment.query.all()
    for es in excel_skills:
        if "Excel Skill" in es.title:
            es.max_marks = 5
            print(f"Updated {es.title} max_marks to 5")
    db.session.commit()
    print("Database updated successfully: Excel skills max_marks set to 5.")
