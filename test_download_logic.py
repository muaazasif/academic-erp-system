from app import app
from excel_assignment import create_excel_exercise_workbook
from io import BytesIO

with app.app_context():
    try:
        title = "Excel Skill 5: VLOOKUP, SUMIF, COUNTIF & IF Formula"
        print(f"Testing title: {title}")
        wb = create_excel_exercise_workbook(assignment_title=title)
        print("Workbook created successfully.")
        
        output = BytesIO()
        wb.save(output)
        print("Workbook saved to BytesIO.")
        
    except Exception as e:
        print(f"Error: {e}")
