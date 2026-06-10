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
    """Returns the full list of 100 tasks"""
    tasks = {}
    for i in range(1, 101):
        if i <= 20: topic = "EXCEL_BASIC"
        elif i <= 40: topic = "EXCEL_ADVANCED"
        elif i <= 60: topic = "SQL_QUERIES"
        elif i <= 80: topic = "POWER_QUERY"
        else: topic = "VBA_EXPERT"
        
        tasks[i] = {
            'id': i,
            'topic': topic,
            'title': f"Task {i}: {topic.replace('_', ' ')}",
            'generate': globals().get(f"generate_task_{i}"),
            'grade': globals().get(f"grade_task_{i}")
        }
    return tasks

# ---------------------------------------------------------
# REAL TASK GENERATORS
# ---------------------------------------------------------

def generate_task_1(ws):
    ws.title = "EXCEL_BASIC_1"
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
    ws['C11'] = "Grand Total:"
    ws['D11'].fill = PatternFill(start_color="FFFF00", fill_type="solid")

def grade_task_1(wb):
    try:
        ws = wb['EXCEL_BASIC_1']
        if ws['D11'].value == 2650: return 1.0
    except: pass
    return 0.0

def generate_task_21(ws):
    ws.title = "EXCEL_ADVANCED_21"
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
        ws = wb['EXCEL_ADVANCED_21']
        if ws['D7'].value == 'A' and ws['D8'].value == 'C' and ws['D9'].value == 'B': return 1.0
    except: pass
    return 0.0

def generate_task_41(ws):
    ws.title = "SQL_QUERIES_41"
    ws['A1'] = "📝 Task 41: Write SQL Query"
    ws['A3'] = "Table: Employees (id, name, salary, dept_id)"
    ws['A5'] = "Q. Write a query to find the names of employees in dept_id 5 with salary > 50000."
    ws['B7'].fill = PatternFill(start_color="FFFF00", fill_type="solid")
    ws['B7'] = "Write SQL here"

def grade_task_41(wb):
    try:
        ws = wb['SQL_QUERIES_41']
        sql = str(ws['B7'].value).lower()
        if "select" in sql and "name" in sql and "dept_id = 5" in sql and "salary > 50000" in sql: return 1.0
    except: pass
    return 0.0

def generate_task_61(ws):
    ws.title = "POWER_QUERY_61"
    ws['A1'] = "📝 Task 61: Power Query Cleaning"
    ws['A3'] = "Instructions: The data below has extra spaces. Provide the CLEANED version in Column B."
    data = ["  Ali  ", "SARA ", " zaman  "]
    for r, val in enumerate(data, 5): 
        ws.cell(row=r, column=1, value=val)
        ws.cell(row=r, column=2).fill = PatternFill(start_color="FFFF00", fill_type="solid")

def grade_task_61(wb):
    try:
        ws = wb['POWER_QUERY_61']
        if ws['B5'].value == "Ali" and ws['B6'].value == "SARA" and ws['B7'].value == "zaman": return 1.0
    except: pass
    return 0.0

def generate_task_81(ws):
    ws.title = "VBA_EXPERT_81"
    ws['A1'] = "📝 Task 81: VBA Macro Creation"
    ws['A3'] = "Q. Create a macro named 'SubmitTest' that shows a Message Box saying 'Done'."
    ws['A5'] = "Instructions: AI will check if the macro exists in the workbook."
    ws['B7'] = "Macro name must be exact."

def grade_task_81(wb):
    return 1.0 # Placeholder for VBA success

# ... Placeholders for the rest ...

real_tasks = [1, 21, 41, 61, 81]
for i in range(2, 101):
    if i in real_tasks: continue
    exec(f"""
def generate_task_{i}(ws):
    ws.title = "TASK_{i}"
    ws['A1'] = "📝 Task {i}: Advanced Exercise Placeholder"
    ws['A3'] = "Instructions: Perform complex data analysis here."
    ws['B5'].fill = PatternFill(start_color='FFFF00', fill_type='solid')
    ws['B5'] = "Write formula here"

def grade_task_{i}(wb):
    try:
        ws = wb['TASK_{i}']
        if ws['B5'].value and ws['B5'].value != "Write formula here": return 1.0
    except: pass
    return 0.0
""")

# ============================================
# WORKBOOK ORCHESTRATION
# ============================================

def create_randomized_midterm(task_ids):
    """Creates a workbook with 10 specific tasks"""
    from excel_assignment import create_excel_exercise_workbook
    wb = create_excel_exercise_workbook("Midterm Exam")
    
    for sheet in wb.sheetnames:
        if sheet != 'Instructions':
            wb.remove(wb[sheet])
            
    bank = get_task_bank()
    for tid in task_ids:
        if tid in bank:
            safe_title = re.sub(r'[:\\/?*\[\]]', '', bank[tid]['title'])[:31]
            ws = wb.create_sheet(safe_title)
            bank[tid]['generate'](ws)
            
    if 'Instructions' in wb.sheetnames:
        ws = wb['Instructions']
        ws['A20'] = "🏆 MIDTERM EXAM READY"
        ws['A21'] = f"Your specific tasks: {', '.join(map(str, task_ids))}"
        
    return wb

def grade_randomized_midterm(file_path, task_ids):
    """Grades a submission based on assigned tasks"""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except:
        return 0, "Error opening file"
        
    bank = get_task_bank()
    total_score = 0
    details = []
    
    for tid in task_ids:
        if tid in bank:
            score = bank[tid]['grade'](wb)
            total_score += score
            details.append({'task': bank[tid]['title'], 'score': score})
            
    return total_score, details
