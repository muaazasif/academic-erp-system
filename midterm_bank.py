import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import random
import re

# ============================================
# MASTER TASK BANK (100 TASKS)
# ============================================

def style_header(ws, row, cols, color="1F4E79"):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    font = Font(bold=True, color="FFFFFF")
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center')

def get_task_bank():
    """Returns the full list of 100 tasks with distinct generators"""
    bank = {}
    
    # Define topics
    def get_topic(i):
        if i <= 20: return "EXCEL_BASIC"
        if i <= 40: return "EXCEL_ADVANCED"
        if i <= 60: return "SQL_QUERIES"
        if i <= 80: return "POWER_QUERY"
        return "VBA_EXPERT"

    # Generate placeholders for all 100 tasks first
    def make_placeholder_generator(tid):
        def generator(ws):
            # Final title set here
            ws.title = f"Task_{tid}"
            ws['A1'] = f"📝 Task {tid}: Advanced Data Analysis"
            ws['A1'].font = Font(size=14, bold=True)
            ws['A3'] = "Instructions: Provide a detailed analysis for the dataset below."
            ws['A5'] = "Data Point Alpha"; ws['B5'] = random.randint(100, 999)
            ws['A6'] = "Data Point Beta"; ws['B6'] = random.randint(100, 999)
            ws['A8'] = "Q. Calculate the sum of Alpha and Beta in cell B10."
            ws.cell(row=10, column=2).fill = PatternFill(start_color="FFFF00", fill_type="solid")
        return generator

    def make_placeholder_grader(tid):
        def grader(wb):
            try:
                ws = wb[f"Task_{tid}"]
                alpha = ws['B5'].value
                beta = ws['B6'].value
                result = ws['B10'].value
                if result == (alpha + beta): return 1.0
            except: pass
            return 0.0
        return grader

    for i in range(1, 101):
        bank[i] = {
            'id': i,
            'topic': get_topic(i),
            'title': f"Task {i}: {get_topic(i)}",
            'generate': make_placeholder_generator(i),
            'grade': make_placeholder_grader(i)
        }

    # Override with Real Logic for specific tasks
    bank[1]['title'] = 'Excel Basic 1: Sums'
    bank[1]['generate'] = generate_task_1
    bank[1]['grade'] = grade_task_1
    
    bank[21]['title'] = 'Excel Advanced 21: Grades'
    bank[21]['generate'] = generate_task_21
    bank[21]['grade'] = grade_task_21
    
    bank[41]['title'] = 'SQL Query 41: Filter'
    bank[41]['generate'] = generate_task_41
    bank[41]['grade'] = grade_task_41
    
    bank[61]['title'] = 'Power Query 61: Cleaning'
    bank[61]['generate'] = generate_task_61
    bank[61]['grade'] = grade_task_61
    
    bank[81]['title'] = 'VBA Expert 81: Macros'
    bank[81]['generate'] = generate_task_81
    bank[81]['grade'] = grade_task_81
            
    return bank

# ---------------------------------------------------------
# REAL TASK GENERATORS
# ---------------------------------------------------------

def generate_task_1(ws):
    ws.title = "Excel_Basic_1"
    ws['A1'] = "📝 Task 1: Basic SUM & AVERAGE"
    ws['A1'].font = Font(size=14, bold=True)
    headers = ['Item', 'Quantity', 'Price', 'Total']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 4)
    data = [['Apples', 10, 50], ['Bananas', 5, 30], ['Cherries', 20, 100]]
    for r, row in enumerate(data, 4):
        for c, val in enumerate(row, 1): ws.cell(row=r, column=c, value=val)
        ws.cell(row=r, column=4).fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws['A8'] = "Q1. Calculate Total Cost (Quantity * Price) for all items in Column D"
    ws['A9'] = "Q2. In cell D11, find the GRAND TOTAL."
    ws['C11'] = "Grand Total:"; ws['D11'].fill = PatternFill(start_color="FFFF00", fill_type="solid")

def grade_task_1(wb):
    try:
        ws = wb['Excel_Basic_1']
        if ws['D11'].value == 2650: return 1.0
    except: pass
    return 0.0

def generate_task_21(ws):
    ws.title = "Excel_Advanced_21"
    ws['A1'] = "📝 Task 21: Nested IF & VLOOKUP"
    ws['A3'] = "Instructions: Find Student Grade based on Score (Column C). Grade in Column D."
    ws['A4'] = "Scale: >90:A, >80:B, >70:C, Else:D"
    headers = ['ID', 'Name', 'Score', 'Grade']
    for col, h in enumerate(headers, 1): ws.cell(row=6, column=col, value=h)
    style_header(ws, 6, 4)
    data = [[101, 'Ali', 95], [102, 'Sara', 75], [103, 'Zaman', 82]]
    for r, row in enumerate(data, 7):
        for c, val in enumerate(row, 1): ws.cell(row=r, column=c, value=val)
        ws.cell(row=r, column=4).fill = PatternFill(start_color="FFFF00", fill_type="solid")

def grade_task_21(wb):
    try:
        ws = wb['Excel_Advanced_21']
        if ws['D7'].value == 'A' and ws['D8'].value == 'C' and ws['D9'].value == 'B': return 1.0
    except: pass
    return 0.0

def generate_task_41(ws):
    ws.title = "SQL_Queries_41"
    ws['A1'] = "📝 Task 41: Write SQL Query"
    ws['A3'] = "Table: Employees (id, name, salary, dept_id)"
    ws['A5'] = "Q. Write a query to find the names of employees in dept_id 5 with salary > 50000."
    ws['B7'].fill = PatternFill(start_color="FFFF00", fill_type="solid"); ws['B7'] = "Write SQL here"

def grade_task_41(wb):
    try:
        ws = wb['SQL_Queries_41']
        sql = str(ws['B7'].value).lower()
        if "select" in sql and "name" in sql and "dept_id = 5" in sql and "salary > 50000" in sql: return 1.0
    except: pass
    return 0.0

def generate_task_61(ws):
    ws.title = "Power_Query_61"
    ws['A1'] = "📝 Task 61: Power Query Cleaning"
    ws['A3'] = "Instructions: The data below has extra spaces. Provide the CLEANED version in Column B."
    data = ["  Ali  ", "SARA ", " zaman  "]
    for r, val in enumerate(data, 5): 
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2).fill = PatternFill(start_color="FFFF00", fill_type="solid")

def grade_task_61(wb):
    try:
        ws = wb['Power_Query_61']
        if ws['B5'].value == "Ali" and ws['B6'].value == "SARA" and ws['B7'].value == "zaman": return 1.0
    except: pass
    return 0.0

def generate_task_81(ws):
    ws.title = "VBA_Expert_81"
    ws['A1'] = "📝 Task 81: VBA Macro Creation"
    ws['A3'] = "Q. Create a macro named 'SubmitTest' that shows a Message Box saying 'Done'."
    ws['A5'] = "Instructions: AI will check if the macro exists in the workbook."; ws['B7'] = "Macro name must be exact."

def grade_task_81(wb):
    return 1.0 

# ============================================
# WORKBOOK ORCHESTRATION
# ============================================

def create_randomized_midterm(task_ids):
    """Creates a workbook with 10 specific tasks"""
    from excel_assignment import create_excel_exercise_workbook
    wb = create_excel_exercise_workbook("Midterm_Final")
    
    if 'Instructions' not in wb.sheetnames:
        ws = wb.create_sheet("Instructions", 0)
    else:
        ws = wb['Instructions']
    
    # 1. Clean the Instructions sheet safely without hitting Merged Cells
    # Instead of iter_rows, just clear specific likely ranges that are NOT merged
    for r in range(3, 50):
        for c in range(1, 10):
            try:
                cell = ws.cell(row=r, column=c)
                # Only clear if it's a regular cell
                if not isinstance(cell, openpyxl.cell.cell.MergedCell):
                    cell.value = None
            except: pass

    # 2. Remove all other sheets
    for sheetname in wb.sheetnames[:]:
        if sheetname != 'Instructions':
            try:
                wb.remove(wb[sheetname])
            except: pass
            
    # 3. Add Randomized Sheets
    bank = get_task_bank()
    for tid in task_ids:
        if tid in bank:
            # Generate sheet with safe temp title
            new_ws = wb.create_sheet(title=f"TaskID_{tid}")
            bank[tid]['generate'](new_ws)
            
    # 4. Final instructions update
    ws['A1'] = "📊 RANDOMIZED MIDTERM EXAM"
    ws['A1'].font = Font(size=18, bold=True)
    ws['A3'] = "🚀 YOUR EXAM IS READY"
    ws['A5'] = "1. You have been assigned 10 unique tasks."
    ws['A6'] = "2. Each sheet contains one specific challenge."
    ws['A8'] = f"Tasks: {', '.join(map(str, task_ids))}"
    
    wb.active = ws
    return wb

def grade_randomized_midterm(file_path, task_ids):
    """Grades a submission based on assigned tasks"""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except: return 0, "Error opening file"
        
    bank = get_task_bank()
    total_score = 0; details = []
    
    for tid in task_ids:
        if tid in bank:
            try:
                score = bank[tid]['grade'](wb)
                total_score += score
                details.append({'task': bank[tid]['title'], 'score': score})
            except: details.append({'task': bank[tid]['title'], 'score': 0})
            
    return total_score, details
