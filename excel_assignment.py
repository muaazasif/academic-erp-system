"""
Excel Skills Assignment Generator & Auto-Grader
With ANTI-CHEATING: Opens other windows = ZERO marks
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import json
import os
import shutil

def style_header(ws, row_num, cols):
    """Style header row"""
    fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    font = Font(bold=True, color="FFFFFF", size=11)
    for c in range(1, cols+1):
        cell = ws.cell(row=row_num, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

def create_excel_exercise_workbook(assignment_title=""):
    """Create workbook with anti-cheating protection"""
    template_path = os.path.join(os.path.dirname(__file__), 'excel_template.xlsm')
    
    if os.path.exists(template_path):
        # Use template with VBA cheating detection
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsm', dir=os.path.dirname(__file__))
        shutil.copy(template_path, temp_file.name)
        temp_file.close()
        
        # MUST use keep_vba=True to preserve macros!
        wb = openpyxl.load_workbook(temp_file.name, keep_vba=True)
        os.unlink(temp_file.name)
    else:
        # Create workbook without VBA (fallback)
        wb = openpyxl.Workbook()
    
    # Remove default sheet if exists
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Determine which exercises to create based on assignment title
    if "Data Validation" in assignment_title and "Manager" in assignment_title:
        create_instructions_dv(wb)
        create_named_manager_exercises(wb)
        create_dropdown_basic_exercises(wb)
        create_dropdown_advanced_exercises(wb)
        create_workbook_structure_exercise(wb)
    elif "Data Cleaning" in assignment_title and "Power Query" in assignment_title:
        create_instructions_skill3(wb)
        create_data_cleaning_exercises(wb)
        create_power_query_exercises(wb)
    else:
        # Default Full Course Workbook (Skill 1)
        create_instructions(wb)
        create_vlookup_exercises(wb)
        create_sumif_countif_exercises(wb)
        create_text_functions_exercises(wb)
        create_if_nested_exercises(wb)
        create_complex_challenge(wb)
    
    # Force macro enablement: Hide all sheets except Instructions
    # ONLY do this if we are using the template (which has the VBA to unhide them)
    if os.path.exists(template_path):
        for sheetname in wb.sheetnames:
            if sheetname != 'Instructions':
                wb[sheetname].sheet_state = 'veryHidden'
    
    # Make Instructions the active sheet
    if 'Instructions' in wb.sheetnames:
        wb.active = wb['Instructions']
    
    return wb

def create_instructions_dv(wb):
    # Create Instructions for Data Validation Assignment
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS: DATA VALIDATION & NAME MANAGER"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = 40
    
    ws['A2'] = "⚠️ IMPORTANT: YOU MUST ENABLE MACROS TO START"
    ws['A2'].font = Font(size=14, bold=True, color="FF0000")
    ws.merge_cells('A2:F2')
    
    ws['A4'] = "🚀 STEPS TO START:"
    ws['A5'] = "1. Enable Content/Macros to see all sheets."
    ws['A6'] = "2. Complete all tasks in YELLOW cells."
    ws['A7'] = "3. AI will grade your submission based on correct validation and named ranges."
    
    ws['A9'] = "📋 ASSIGNMENT MODULES:"
    ws['A10'] = "1. Named Manager (2.5 marks)"
    ws['A11'] = "2. Dropdown Basic (2.5 marks)"
    ws['A12'] = "3. Dropdown Advanced (2.5 marks)"
    ws['A13'] = "4. Workbook Data Validation (2.5 marks)"
    
    ws.column_dimensions['A'].width = 55

def create_named_manager_exercises(wb):
    ws = wb.create_sheet("NAMED MANAGER")
    ws['A1'] = "📝 Task: Create Named Ranges"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    ws['A3'] = "1. Select cells G4:G8 and name them 'ProductList'"
    ws['A4'] = "2. Select cells H4:H8 and name them 'PriceList'"
    ws['A5'] = "3. Select cells I4:I8 and name them 'CategoryList'"
    
    headers = ['Products', 'Prices', 'Categories']
    for col, h in enumerate(headers, 7):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 3)
    
    data = [
        ['Laptop', 50000, 'Electronics'],
        ['Mouse', 1500, 'Accessories'],
        ['Keyboard', 3500, 'Accessories'],
        ['Monitor', 15000, 'Electronics'],
        ['USB', 1200, 'Storage']
    ]
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=7+j, value=val)

def create_dropdown_basic_exercises(wb):
    ws = wb.create_sheet("DROPDOWN BASIC")
    ws['A1'] = "📝 Task: Basic Dropdowns"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    ws['A3'] = "1. In cell C5, create a dropdown using the list: Apple, Mango, Banana, Orange"
    ws['A4'] = "2. In cell C7, create a dropdown using the Named Range 'ProductList' from the previous sheet"
    
    ws['B5'] = "Select Fruit:"
    ws['C5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B7'] = "Select Product:"
    ws['C7'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20

def create_dropdown_advanced_exercises(wb):
    ws = wb.create_sheet("DROPDOWN ADVANCED")
    ws['A1'] = "📝 Task: Dependent Dropdowns"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    ws['A3'] = "1. Create Named Ranges for 'Electronics' (Laptop, Mobile) and 'Furniture' (Chair, Table)"
    ws['A4'] = "2. In cell C10, create a dropdown for Category (Electronics, Furniture)"
    ws['A5'] = "3. In cell D10, create a DEPENDENT dropdown that shows items based on C10"
    
    # Data for naming
    ws['G3'] = "Electronics"
    ws['G4'] = "Laptop"
    ws['G5'] = "Mobile"
    
    ws['H3'] = "Furniture"
    ws['H4'] = "Chair"
    ws['H5'] = "Table"
    
    ws['B10'] = "Category:"
    ws['C10'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B11'] = "Item:"
    ws['D10'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def create_workbook_structure_exercise(wb):
    ws = wb.create_sheet("DATA VALIDATION SKILL 2")
    ws['A1'] = "📝 Task: Whole Number & Date Validation"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    ws['A3'] = "1. Apply validation to cell C5: Whole Number between 10 and 100"
    ws['A4'] = "2. Apply validation to cell C7: Date between 2024-01-01 and 2024-12-31"
    ws['A5'] = "3. Apply validation to cell C9: Text Length exactly 5 characters"
    
    ws['B5'] = "Enter Number (10-100):"
    ws['C5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws['B7'] = "Enter Date (2024):"
    ws['C7'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws['B9'] = "Enter Code (5 chars):"
    ws['C9'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30

def create_instructions_skill3(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS: DATA CLEANING & POWER QUERY"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = 40
    
    ws['A3'] = "🚀 STEPS TO START:"
    ws['A4'] = "1. Enable Macros to see all task sheets."
    ws['A5'] = "2. Task 1: Clean raw data using functions (TRIM, PROPER, etc.)."
    ws['A6'] = "3. Task 2: Use Flash Fill or Text-to-Columns."
    ws['A7'] = "4. Task 3: Handle Duplicates and Power Query transformation."
    
    ws['A9'] = "📋 ASSIGNMENT MODULES (Total 10 Marks):"
    ws['A10'] = "1. Basic Data Cleaning (5 marks)"
    ws['A11'] = "2. Power Query & Transformations (5 marks)"
    
    ws.column_dimensions['A'].width = 55

def create_data_cleaning_exercises(wb):
    ws = wb.create_sheet("DATA CLEANING")
    ws['A1'] = "📝 Task 1: Data Cleaning Functions (TRIM, PROPER, UPPER)"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    headers = ['Raw Data', 'Clean Data (Expected Formula)']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 2)
    
    raw_data = [
        "   muaaz asif   ",
        "PYTHON PROGRAMMING",
        "excel   skills",
        "john doe  ",
        "   KARAchi-PAKistan"
    ]
    
    for i, data in enumerate(raw_data):
        ws.cell(row=4+i, column=1, value=data)
        ws.cell(row=4+i, column=2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws['A11'] = "📝 Task 2: Flash Fill / Text-to-Columns"
    ws['A12'] = "Separate 'First Name' and 'Last Name' from the Full Name column."
    
    ws['A14'] = "Full Name"
    ws['B14'] = "First Name"
    ws['C14'] = "Last Name"
    style_header(ws, 14, 3)
    
    names = ["Ali Khan", "Sara Ahmed", "Bilal Sheikh"]
    for i, name in enumerate(names):
        ws.cell(row=15+i, column=1, value=name)
        ws.cell(row=15+i, column=2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws.cell(row=15+i, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30

def create_power_query_exercises(wb):
    ws = wb.create_sheet("POWER QUERY")
    ws['A1'] = "📝 Task 3: Data Transformation & Duplicates"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    ws['A3'] = "1. Identify and remove duplicate rows from the table below (F4:H10)."
    ws['A4'] = "2. Ensure all text is in UPPERCASE in the 'Product' column."
    
    data = [
        ['ID', 'Product', 'Sales'],
        [101, 'Laptop', 500],
        [102, 'Mouse', 50],
        [101, 'Laptop', 500], # Duplicate
        [103, 'Keyboard', 80],
        [102, 'Mouse', 50], # Duplicate
        [104, 'Monitor', 300]
    ]
    
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=6+j, value=val)
            if i > 0:
                ws.cell(row=4+i, column=6+j).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    style_header(ws, 4, 3) # Note: headers start at F4
    
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 15

def create_instructions(wb):
    # Create Instructions as the first sheet
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS ASSIGNMENT - Complete Workbook"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = 40
    
    ws['A2'] = "⚠️ IMPORTANT: YOU MUST ENABLE MACROS TO START"
    ws['A2'].font = Font(size=14, bold=True, color="FF0000")
    ws.merge_cells('A2:F2')
    
    ws['A3'] = "🚀 STEPS TO START:"
    ws['A3'].font = Font(bold=True)
    ws['A4'] = "1. If you see a 'Security Risk' bar: Close Excel -> Right-click this file -> Properties -> Check 'Unblock' -> OK."
    ws['A5'] = "2. Open file again and click 'Enable Content' or 'Enable Macros'."
    ws['A6'] = "3. All exercise sheets will appear automatically once macros are enabled."
    
    ws['A8'] = "🚨 ANTI-CHEATING: IF YOU OPEN ANY OTHER WINDOW, YOUR MARKS WILL BE ZERO!"
    ws['A8'].font = Font(size=11, bold=True, color="C00000")
    ws.merge_cells('A8:F8')
    
    content = [
        ("", None),
        ("📋 OVERVIEW:", None),
        ("• Total Marks: 10", None),
        ("• Total Questions: 30+ (Easy → Hard)", None),
        ("• Time Limit: 2 hours", None),
        ("", None),
        ("📝 EXERCISE SHEETS:", None),
        ("1. VLOOKUP (2 marks) - 4 questions", None),
        ("2. SUMIF & COUNTIF (2 marks) - 6 questions", None),
        ("3. LEFT, RIGHT, MID (2 marks) - 6 questions", None),
        ("4. IF & NESTED IF (2 marks) - 10 questions (Graded per student)", None),
        ("5. COMPLEX CHALLENGE (2 marks) - 10 questions", None),
        ("", None),
        ("✅ HOW TO COMPLETE:", None),
        ("• Write formula in YELLOW cell", None),
        ("• Your answer calculates automatically", None),
        ("• DO NOT change sheet names", None),
        ("• DO NOT switch windows or search on Google", None)
    ]
    
    for i, (text, _) in enumerate(content, 5):
        ws.cell(row=i, column=1, value=text)
    
    ws.column_dimensions['A'].width = 55

def create_vlookup_exercises(wb):
    ws = wb.create_sheet("VLOOKUP")
    
    ws.merge_cells('A1:E1')
    ws['A1'] = "📊 Employee Database"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    # Sample data
    headers = ['Emp ID', 'Name', 'Department', 'City', 'Salary']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 5)
    
    data = [
        ['E001', 'Ahmed Ali', 'IT', 'Karachi', 45000],
        ['E002', 'Sara Khan', 'HR', 'Lahore', 42000],
        ['E003', 'Omar Sheikh', 'Finance', 'Islamabad', 50000],
        ['E004', 'Fatima Noor', 'IT', 'Karachi', 48000],
        ['E005', 'Bilal Ahmed', 'Sales', 'Peshawar', 38000],
        ['E006', 'Ayesha Malik', 'HR', 'Lahore', 41000],
        ['E007', 'Hassan Raza', 'Finance', 'Islamabad', 52000],
        ['E008', 'Zainab Hussain', 'Sales', 'Quetta', 36000],
        ['E009', 'Ali Raza', 'IT', 'Karachi', 46000],
        ['E010', 'Maryam Fatima', 'HR', 'Lahore', 43000],
    ]
    
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 16
    ws.merge_cells(f'A{q_row}:F{q_row}')
    ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use VLOOKUP formula").font = Font(size=12, bold=True, color="C00000")
    
    ws.cell(row=q_row+2, column=1, value="Q#").font = Font(bold=True)
    ws.cell(row=q_row+2, column=2, value="Question").font = Font(bold=True)
    ws.cell(row=q_row+2, column=3, value="Your Formula (YELLOW cell)").font = Font(bold=True)
    ws.cell(row=q_row+2, column=4, value="Your Answer").font = Font(bold=True)
    style_header(ws, q_row+2, 4)
    
    questions = [
        ['Q1', 'Find Department of Employee E003'],
        ['Q2', 'Find Salary of Sara Khan'],
        ['Q3', 'Find City of Employee E007'],
        ['Q4', 'Find Name of Employee with ID E010'],
    ]
    
    for i, q in enumerate(questions):
        r = q_row + 3 + i
        ws.cell(row=r, column=1, value=q[0])
        ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3, value="").fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 20

def create_sumif_countif_exercises(wb):
    ws = wb.create_sheet("SUMIF & COUNTIF")
    
    ws.merge_cells('A1:F1')
    ws['A1'] = "📊 Sales Data - Q1 2024"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    headers = ['Date', 'Salesperson', 'Product', 'Category', 'Quantity', 'Amount']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 6)
    
    data = [
        ['2024-01-05', 'Ali', 'Laptop', 'Electronics', 2, 120000],
        ['2024-01-10', 'Sara', 'Mouse', 'Accessories', 10, 15000],
        ['2024-01-15', 'Ali', 'Keyboard', 'Accessories', 5, 25000],
        ['2024-02-01', 'Omar', 'Laptop', 'Electronics', 3, 180000],
        ['2024-02-10', 'Sara', 'Monitor', 'Electronics', 4, 80000],
        ['2024-02-15', 'Ali', 'Mouse', 'Accessories', 8, 12000],
        ['2024-03-01', 'Omar', 'Keyboard', 'Accessories', 6, 30000],
        ['2024-03-10', 'Fatima', 'Laptop', 'Electronics', 2, 120000],
        ['2024-03-15', 'Sara', 'Laptop', 'Electronics', 1, 60000],
        ['2024-03-20', 'Fatima', 'Monitor', 'Electronics', 3, 60000],
    ]
    
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 17
    ws.merge_cells(f'A{q_row}:F{q_row}')
    ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use SUMIF and COUNTIF").font = Font(size=12, bold=True, color="C00000")
    
    ws.cell(row=q_row+2, column=1, value="Q#").font = Font(bold=True)
    ws.cell(row=q_row+2, column=2, value="Question").font = Font(bold=True)
    ws.cell(row=q_row+2, column=3, value="Formula").font = Font(bold=True)
    ws.cell(row=q_row+2, column=4, value="Answer").font = Font(bold=True)
    style_header(ws, q_row+2, 4)
    
    questions = [
        ['Q5', 'Total sales by Ali (SUMIF)'],
        ['Q6', 'Count of Laptop sales (COUNTIF)'],
        ['Q7', 'Total quantity sold by Sara (SUMIF)'],
        ['Q8', 'Count of Electronics sold (COUNTIF)'],
        ['Q9', 'Total amount of Accessories (SUMIF)'],
        ['Q10', 'Count of Fatima sales (COUNTIF)'],
    ]
    
    for i, q in enumerate(questions):
        r = q_row + 3 + i
        ws.cell(row=r, column=1, value=q[0])
        ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3, value="").fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 20

def create_text_functions_exercises(wb):
    ws = wb.create_sheet("LEFT RIGHT MID")
    
    ws.merge_cells('A1:D1')
    ws['A1'] = "📊 Student Data"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    headers = ['Full Name', 'Phone Number', 'Email', 'Student Code']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 4)
    
    data = [
        ['Ahmed Ali Khan', '0300-1234567', 'ahmed.ali@gmail.com', 'STU-2024-001'],
        ['Sara Fatima', '0321-7654321', 'sara.f@hotmail.com', 'STU-2024-002'],
        ['Muhammad Omar', '0333-9876543', 'omar.pk@yahoo.com', 'STU-2024-003'],
        ['Fatima Noor', '0345-1112233', 'fatima.noor@outlook.com', 'STU-2024-004'],
        ['Bilal Hussain', '0301-4445566', 'bilal.h@gmail.com', 'STU-2024-005'],
    ]
    
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 12
    ws.merge_cells(f'A{q_row}:E{q_row}')
    ws.cell(row=q_row, column=1, value="📝 EXERCISES - LEFT, RIGHT, MID, LEN, FIND").font = Font(size=12, bold=True, color="C00000")
    
    ws.cell(row=q_row+2, column=1, value="Q#").font = Font(bold=True)
    ws.cell(row=q_row+2, column=2, value="Question").font = Font(bold=True)
    ws.cell(row=q_row+2, column=3, value="Formula").font = Font(bold=True)
    ws.cell(row=q_row+2, column=4, value="Answer").font = Font(bold=True)
    style_header(ws, q_row+2, 4)
    
    questions = [
        ['Q11', 'Extract first 3 letters from A4 using LEFT'],
        ['Q12', 'Extract last 7 digits from B4 using RIGHT'],
        ['Q13', 'Extract username from C4 (before @) using LEFT + FIND'],
        ['Q14', 'Extract year from D4 (2024) using MID'],
        ['Q15', 'Extract domain from C5 (after @) using RIGHT + LEN + FIND'],
        ['Q16', 'Extract middle name from A5 using MID + FIND'],
    ]
    
    for i, q in enumerate(questions):
        r = q_row + 3 + i
        ws.cell(row=r, column=1, value=q[0])
        ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3, value="").fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 25

def create_if_nested_exercises(wb):
    ws = wb.create_sheet("IF & NESTED IF")

    ws.merge_cells('A1:E1')
    ws['A1'] = "📊 Student Marks Data"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")

    headers = ['Student ID', 'Name', 'Marks', 'Grade (Use IF)', 'Status (Use IF)']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 5)

    data = [
        ['S001', 'Ahmed', 92, '', ''],
        ['S002', 'Sara', 85, '', ''],
        ['S003', 'Omar', 76, '', ''],
        ['S004', 'Fatima', 65, '', ''],
        ['S005', 'Bilal', 58, '', ''],
        ['S006', 'Ayesha', 45, '', ''],
        ['S007', 'Hassan', 35, '', ''],
        ['S008', 'Zainab', 88, '', ''],
        ['S009', 'Ali', 72, '', ''],
        ['S010', 'Maryam', 95, '', ''],
    ]

    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=1+j, value=val)

    # Grading scale in column G (no merging to avoid read-only cells)
    ws['G3'] = "📝 GRADING SCALE:"
    ws['G3'].font = Font(bold=True, size=11, color="1F4E79")

    grading = [
        "90-100 = A+",
        "80-89 = A",
        "70-79 = B",
        "60-69 = C",
        "50-59 = D",
        "Below 50 = F",
        "",
        ">= 50 = Pass",
        "< 50 = Fail"
    ]

    for i, text in enumerate(grading, 4):
        ws.cell(row=3+i, column=7, value=text)

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['G'].width = 25

def create_data_validation_exercises(wb):
    ws = wb.create_sheet("DATA VALIDATION SKILL 1")
    
    ws.merge_cells('A1:G1')
    ws['A1'] = "📊 Data Validation & Named Ranges"
    ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    
    # Instructions for students
    ws['A3'] = "🚀 INSTRUCTIONS:"
    ws['A4'] = "1. Create a Named Range 'Departments' for cells I4:I8"
    ws['A5'] = "2. Create a Named Range 'Cities' for cells J4:J8"
    ws['A6'] = "3. Apply Data Validation (Whole Number 1-100) to 'Emp ID' column"
    ws['A7'] = "4. Apply List Validation (Dropdown) to 'Dept' column using 'Departments' range"
    ws['A8'] = "5. Apply List Validation (Dropdown) to 'City' column using 'Cities' range"
    ws['A9'] = "6. Create Dependent Dropdown for 'Manager' (Advanced: if IT -> Manager A, if HR -> Manager B)"
    
    # Manager Data for reference
    ws['I3'] = "DEPARTMENTS"
    ws['J3'] = "CITIES"
    ws['K3'] = "MANAGERS"
    style_header(ws, 3, 11) # Header for range area
    
    depts = ['IT', 'HR', 'Finance', 'Sales', 'Admin']
    cities = ['Karachi', 'Lahore', 'Islamabad', 'Quetta', 'Peshawar']
    managers = ['Manager A', 'Manager B', 'Manager C', 'Manager D', 'Manager E']
    
    for i in range(5):
        ws.cell(row=4+i, column=9, value=depts[i])
        ws.cell(row=4+i, column=10, value=cities[i])
        ws.cell(row=4+i, column=11, value=managers[i])

    # Table Area
    headers = ['Emp ID', 'Emp Name', 'Dept', 'City', 'Manager']
    for col, h in enumerate(headers, 1):
        ws.cell(row=12, column=col, value=h)
    style_header(ws, 12, 5)
    
    # Add 5 empty rows for student to fill/apply validation
    for r in range(13, 18):
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            if c >= 3: # Validation targets
                 ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 15
    ws.column_dimensions['K'].width = 15

def create_complex_challenge(wb):
    ws = wb.create_sheet("COMPLEX CHALLENGE")
    
    ws.merge_cells('A1:F1')
    ws['A1'] = "🏆 CHALLENGE - Complete Product Analysis"
    ws['A1'].font = Font(size=14, bold=True, color="C00000")
    
    headers = ['Product Code', 'Product Name', 'Category', 'Price', 'Stock', 'Status']
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 6)
    
    data = [
        ['PRD-LAP-001', 'Laptop Pro 15', 'Electronics', 85000, 12, ''],
        ['PRD-MOU-002', 'Wireless Mouse', 'Accessories', 1500, 50, ''],
        ['PRD-KEY-003', 'Mech Keyboard', 'Accessories', 4500, 25, ''],
        ['PRD-MON-004', 'Monitor 27 inch', 'Electronics', 35000, 8, ''],
        ['PRD-USB-005', 'USB Hub 7-in-1', 'Accessories', 2500, 40, ''],
        ['PRD-HDD-006', 'External HDD 1TB', 'Storage', 8000, 15, ''],
        ['PRD-SSD-007', 'SSD 500GB', 'Storage', 6500, 20, ''],
        ['PRD-WEB-008', 'Webcam HD', 'Accessories', 3500, 30, ''],
        ['PRD-TAB-009', 'Tablet 10 inch', 'Electronics', 45000, 5, ''],
        ['PRD-SPK-010', 'Bluetooth Speaker', 'Accessories', 5500, 35, ''],
    ]
    
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 17
    ws.merge_cells(f'A{q_row}:F{q_row}')
    ws.cell(row=q_row, column=1, value="📝 CHALLENGE - Combine all functions!").font = Font(size=12, bold=True, color="C00000")
    
    ws.cell(row=q_row+2, column=1, value="Q#").font = Font(bold=True)
    ws.cell(row=q_row+2, column=2, value="Question").font = Font(bold=True)
    ws.cell(row=q_row+2, column=3, value="Formula").font = Font(bold=True)
    ws.cell(row=q_row+2, column=4, value="Answer").font = Font(bold=True)
    style_header(ws, q_row+2, 4)
    
    questions = [
        ['Q21', 'Extract category code from A4 (e.g., "LAP") using MID'],
        ['Q22', 'Count products in "Accessories" using COUNTIF'],
        ['Q23', 'Total stock of Electronics using SUMIF'],
        ['Q24', 'Find price of PRD-KEY-003 using VLOOKUP'],
        ['Q25', 'Extract first word from B5 using LEFT + FIND'],
        ['Q26', 'Status: IF stock > 20 then "In Stock" else "Low"'],
        ['Q27', 'Total value of all Accessories (Price × Stock, SUMIF)'],
        ['Q28', 'Extract number from A5 (003) using RIGHT'],
        ['Q29', 'Count products with price > 10000 using COUNTIF'],
        ['Q30', 'Extract domain from email "admin@company.com" using RIGHT + LEN + FIND'],
    ]
    
    for i, q in enumerate(questions):
        r = q_row + 3 + i
        ws.cell(row=r, column=1, value=q[0])
        ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3, value="").fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 55
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 20

# ============================================
# AUTO-GRADER
# ============================================

def grade_excel_submission(file_path, assignment_title=""):
    """Auto-grade a completed Excel assignment. Returns score out of 10"""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False) # data_only=False to see formulas/validations
    except Exception as e:
        return {'error': f'Cannot open file: {str(e)}', 'score': 0}
    
    # ==========================================
    # CHECK FOR CHEATING
    # ==========================================
    cheating_detected = False
    macros_disabled = False
    
    if 'Instructions' in wb.sheetnames:
        ws = wb['Instructions']
        macro_flag = ws.cell(row=99, column=26).value   # Z99
        cheat_flag = ws.cell(row=100, column=26).value  # Z100
        
        if not macro_flag or str(macro_flag).upper() != 'MACROS_OK':
            macros_disabled = True
        
        if cheat_flag and 'CHEAT' in str(cheat_flag).upper():
            cheating_detected = True
    
    if macros_disabled:
        return {
            'score': 0, 'max': 10, 'percentage': 0, 'cheating_detected': False, 'macros_disabled': True,
            'details': {'error': 'Macros were not enabled. You must enable macros to complete the assignment.'}
        }

    if cheating_detected:
        return {
            'score': 0.0, 'max': 10, 'percentage': 0.0, 'cheating_detected': True, 'macros_disabled': False,
            'details': {'error': 'CHEATING DETECTED: You switched windows while completing the assignment.'}
        }
    
    total_score = 0
    details = {}
    
    if "Data Validation" in assignment_title and "Manager" in assignment_title:
        nm_score, nm_detail = grade_named_manager(wb)
        db_score, db_detail = grade_dropdown_basic(wb)
        da_score, da_detail = grade_dropdown_advanced(wb)
        wv_score, wv_detail = grade_workbook_validation(wb)
        
        total_score = nm_score + db_score + da_score + wv_score
        details['Named Manager'] = {'score': nm_score, 'max': 2.5, 'details': nm_detail}
        details['Dropdown Basic'] = {'score': db_score, 'max': 2.5, 'details': db_detail}
        details['Dropdown Advanced'] = {'score': da_score, 'max': 2.5, 'details': da_detail}
        details['Workbook Validation'] = {'score': wv_score, 'max': 2.5, 'details': wv_detail}
    elif "Data Cleaning" in assignment_title and "Power Query" in assignment_title:
        # Skill 3 Grading
        wb_vals = openpyxl.load_workbook(file_path, data_only=True)
        dc_score, dc_detail = grade_data_cleaning(wb_vals)
        pq_score, pq_detail = grade_power_query(wb_vals)
        
        total_score = dc_score + pq_score
        details['Data Cleaning'] = {'score': dc_score, 'max': 5, 'details': dc_detail}
        details['Power Query Basics'] = {'score': pq_score, 'max': 5, 'details': pq_detail}
    else:
        # Skill 1 Grading
... (around line 1087) ...
    except: pass
    return min(score, 2), details

def grade_data_cleaning(wb):
    score = 0
    details = []
    try:
        ws = wb['DATA CLEANING']
        # Task 1: Clean functions
        expected = ["Muaaz Asif", "Python Programming", "Excel Skills", "John Doe", "Karachi-Pakistan"]
        for i, exp in enumerate(expected):
            val = ws.cell(row=4+i, column=2).value
            if val and str(val).strip().lower() == exp.lower():
                score += 0.6
                details.append({'task': f'Clean: {exp}', 'correct': True})
            else:
                details.append({'task': f'Clean: {exp}', 'correct': False})
        
        # Task 2: Flash Fill (First/Last names)
        names = [("Ali", "Khan"), ("Sara", "Ahmed"), ("Bilal", "Sheikh")]
        for i, (f, l) in enumerate(names):
            fv = ws.cell(row=15+i, column=2).value
            lv = ws.cell(row=15+i, column=3).value
            if fv and str(fv).strip().lower() == f.lower() and lv and str(lv).strip().lower() == l.lower():
                score += 0.66
                details.append({'task': f'Split: {f} {l}', 'correct': True})
            else:
                details.append({'task': f'Split: {f} {l}', 'correct': False})
    except: pass
    return min(score, 5), details

def grade_power_query(wb):
    score = 0
    details = []
    try:
        ws = wb['POWER QUERY']
        # Check if duplicates are gone (Rows 6 and 8 were duplicates)
        # We expect a unique list starting from F5
        # For simplicity, we check if the values in G are cleaned and unique
        products = []
        for r in range(5, 11):
            p = ws.cell(row=r, column=7).value
            if p: products.append(str(p).strip().upper())
        
        if len(products) == 5 and "LAPTOP" in products and "MOUSE" in products:
            score += 2.5
            details.append({'task': 'Duplicates Removed', 'correct': True})
        else:
            details.append({'task': 'Duplicates Removed', 'correct': False, 'error': 'Table should have 5 unique rows'})
            
        # Check if all products are UPPERCASE
        if all(p.isupper() for p in products if p):
            score += 2.5
            details.append({'task': 'Uppercase Transformation', 'correct': True})
        else:
            details.append({'task': 'Uppercase Transformation', 'correct': False})
    except: pass
    return min(score, 5), details

        wb_vals = openpyxl.load_workbook(file_path, data_only=True)
        
        v_score, v_detail = grade_vlookup(wb_vals)
        total_score += v_score
        details['VLOOKUP'] = {'score': v_score, 'max': 2, 'details': v_detail}
        
        s_score, s_detail = grade_sumif_countif(wb_vals)
        total_score += s_score
        details['SUMIF/COUNTIF'] = {'score': s_score, 'max': 2, 'details': s_detail}
        
        t_score, t_detail = grade_text_functions(wb_vals)
        total_score += t_score
        details['Text Functions'] = {'score': t_score, 'max': 2, 'details': t_detail}

        if_score, if_detail = grade_if_nested(wb_vals)
        total_score += if_score
        details['Nested IF'] = {'score': if_score, 'max': 2, 'details': if_detail}
        
        c_score, c_detail = grade_complex(wb_vals)
        total_score += c_score
        details['Complex'] = {'score': c_score, 'max': 2, 'details': c_detail}
    
    return {
        'score': round(min(total_score, 10), 2),
        'max': 10,
        'percentage': round((min(total_score, 10) / 10) * 100, 1),
        'cheating_detected': False,
        'details': details
    }

def grade_named_manager(wb):
    score = 0
    details = []
    try:
        names = wb.defined_names.keys()
        if 'ProductList' in names: score += 0.8; details.append({'task': 'Named Range: ProductList', 'correct': True})
        else: details.append({'task': 'Named Range: ProductList', 'correct': False, 'error': 'Missing ProductList'})
        if 'PriceList' in names: score += 0.8; details.append({'task': 'Named Range: PriceList', 'correct': True})
        else: details.append({'task': 'Named Range: PriceList', 'correct': False, 'error': 'Missing PriceList'})
        if 'CategoryList' in names: score += 0.9; details.append({'task': 'Named Range: CategoryList', 'correct': True})
        else: details.append({'task': 'Named Range: CategoryList', 'correct': False, 'error': 'Missing CategoryList'})
    except: pass
    return min(score, 2.5), details

def grade_dropdown_basic(wb):
    score = 0
    details = []
    try:
        if 'DROPDOWN BASIC' in wb.sheetnames:
            ws = wb['DROPDOWN BASIC']
            validations = ws.data_validations.dataValidation
            c5_val = any('C5' in rng.coord for dv in validations for rng in dv.sqref if dv.type == 'list')
            c7_val = any('C7' in rng.coord for dv in validations for rng in dv.sqref if dv.type == 'list')
            if c5_val: score += 1.25; details.append({'task': 'C5 Dropdown', 'correct': True})
            else: details.append({'task': 'C5 Dropdown', 'correct': False})
            if c7_val: score += 1.25; details.append({'task': 'C7 Dropdown', 'correct': True})
            else: details.append({'task': 'C7 Dropdown', 'correct': False})
    except: pass
    return min(score, 2.5), details

def grade_dropdown_advanced(wb):
    score = 0
    details = []
    try:
        names = wb.defined_names.keys()
        if 'Electronics' in names and 'Furniture' in names: score += 1.0; details.append({'task': 'Category Ranges', 'correct': True})
        else: details.append({'task': 'Category Ranges', 'correct': False})
        if 'DROPDOWN ADVANCED' in wb.sheetnames:
            ws = wb['DROPDOWN ADVANCED']
            validations = ws.data_validations.dataValidation
            c10_val = any('C10' in rng.coord for dv in validations for rng in dv.sqref)
            d10_val = any('D10' in rng.coord for dv in validations for rng in dv.sqref if 'INDIRECT' in str(dv.formula1).upper())
            if c10_val: score += 0.5; details.append({'task': 'C10 Category', 'correct': True})
            else: details.append({'task': 'C10 Category', 'correct': False})
            if d10_val: score += 1.0; details.append({'task': 'D10 Dependent', 'correct': True})
            else: details.append({'task': 'D10 Dependent', 'correct': False})
    except: pass
    return min(score, 2.5), details

def grade_workbook_validation(wb):
    score = 0
    details = []
    try:
        if 'DATA VALIDATION SKILL 2' in wb.sheetnames:
            ws = wb['DATA VALIDATION SKILL 2']
            validations = ws.data_validations.dataValidation
            c5_val = any('C5' in rng.coord for dv in validations for rng in dv.sqref if dv.type == 'whole')
            c7_val = any('C7' in rng.coord for dv in validations for rng in dv.sqref if dv.type == 'date')
            c9_val = any('C9' in rng.coord for dv in validations for rng in dv.sqref if dv.type == 'textLength')
            if c5_val: score += 0.8; details.append({'task': 'C5 Whole Number', 'correct': True})
            else: details.append({'task': 'C5 Whole Number', 'correct': False})
            if c7_val: score += 0.8; details.append({'task': 'C7 Date', 'correct': True})
            else: details.append({'task': 'C7 Date', 'correct': False})
            if c9_val: score += 0.9; details.append({'task': 'C9 Text Length', 'correct': True})
            else: details.append({'task': 'C9 Text Length', 'correct': False})
    except: pass
    return min(score, 2.5), details

def grade_vlookup(wb):
    score = 0
    details = []
    try:
        ws = wb['VLOOKUP']
        if ws['D19'].value and 'finance' in str(ws['D19'].value).lower(): score += 0.5; details.append({'q': 'Q1', 'correct': True})
        else: details.append({'q': 'Q1', 'correct': False})
        if ws['D20'].value and str(ws['D20'].value).strip() in ['42000', '42000.0']: score += 0.5; details.append({'q': 'Q2', 'correct': True})
        else: details.append({'q': 'Q2', 'correct': False})
        if ws['D21'].value and 'islamabad' in str(ws['D21'].value).lower(): score += 0.5; details.append({'q': 'Q3', 'correct': True})
        else: details.append({'q': 'Q3', 'correct': False})
        if ws['D22'].value and 'maryam' in str(ws['D22'].value).lower(): score += 0.5; details.append({'q': 'Q4', 'correct': True})
        else: details.append({'q': 'Q4', 'correct': False})
    except: pass
    return score, details

def grade_sumif_countif(wb):
    score = 0
    details = []
    try:
        ws = wb['SUMIF & COUNTIF']
        if ws['D20'].value and str(ws['D20'].value).strip() in ['157000', '157000.0']: score += 0.33; details.append({'q': 'Q5', 'correct': True})
        else: details.append({'q': 'Q5', 'correct': False})
        if ws['D21'].value and str(ws['D21'].value).strip() in ['4', '4.0']: score += 0.33; details.append({'q': 'Q6', 'correct': True})
        else: details.append({'q': 'Q6', 'correct': False})
        if ws['D22'].value and str(ws['D22'].value).strip() in ['15', '15.0']: score += 0.33; details.append({'q': 'Q7', 'correct': True})
        else: details.append({'q': 'Q7', 'correct': False})
        if ws['D23'].value and str(ws['D23'].value).strip() in ['7', '7.0']: score += 0.33; details.append({'q': 'Q8', 'correct': True})
        else: details.append({'q': 'Q8', 'correct': False})
        if ws['D24'].value and str(ws['D24'].value).strip() in ['82000', '82000.0']: score += 0.33; details.append({'q': 'Q9', 'correct': True})
        else: details.append({'q': 'Q9', 'correct': False})
        if ws['D25'].value and str(ws['D25'].value).strip() in ['2', '2.0']: score += 0.33; details.append({'q': 'Q10', 'correct': True})
        else: details.append({'q': 'Q10', 'correct': False})
    except: pass
    return min(score, 2), details

def grade_text_functions(wb):
    score = 0
    details = []
    try:
        ws = wb['LEFT RIGHT MID']
        if ws['D15'].value and str(ws['D15'].value).strip().lower() == 'ahm': score += 0.33; details.append({'q': 'Q11', 'correct': True})
        else: details.append({'q': 'Q11', 'correct': False})
        if ws['D16'].value and '1234567' in str(ws['D16'].value): score += 0.33; details.append({'q': 'Q12', 'correct': True})
        else: details.append({'q': 'Q12', 'correct': False})
        if ws['D17'].value and 'ahmed' in str(ws['D17'].value).lower(): score += 0.33; details.append({'q': 'Q13', 'correct': True})
        else: details.append({'q': 'Q13', 'correct': False})
        if ws['D18'].value and '2024' in str(ws['D18'].value): score += 0.33; details.append({'q': 'Q14', 'correct': True})
        else: details.append({'q': 'Q14', 'correct': False})
        if ws['D19'].value and 'hotmail' in str(ws['D19'].value).lower(): score += 0.33; details.append({'q': 'Q15', 'correct': True})
        else: details.append({'q': 'Q15', 'correct': False})
        if ws['D20'].value and 'fatima' in str(ws['D20'].value).lower(): score += 0.33; details.append({'q': 'Q16', 'correct': True})
        else: details.append({'q': 'Q16', 'correct': False})
    except: pass
    return min(score, 2), details

def grade_if_nested(wb):
    score = 0
    details = []
    try:
        if 'IF & NESTED IF' not in wb.sheetnames:
            return 0, [{'task': 'Sheet Missing', 'correct': False, 'error': 'IF & NESTED IF sheet missing'}]
            
        ws = wb['IF & NESTED IF']
        students = [
            {'row': 4, 'name': 'Ahmed', 'g': 'A+', 's': 'Pass'},
            {'row': 5, 'name': 'Sara', 'g': 'A', 's': 'Pass'},
            {'row': 6, 'name': 'Omar', 'g': 'B', 's': 'Pass'},
            {'row': 7, 'name': 'Fatima', 'g': 'C', 's': 'Pass'},
            {'row': 8, 'name': 'Bilal', 'g': 'D', 's': 'Pass'},
            {'row': 9, 'name': 'Ayesha', 'g': 'F', 's': 'Fail'},
            {'row': 10, 'name': 'Hassan', 'g': 'F', 's': 'Fail'},
            {'row': 11, 'name': 'Zainab', 'g': 'A', 's': 'Pass'},
            {'row': 12, 'name': 'Ali', 'g': 'B', 's': 'Pass'},
            {'row': 13, 'name': 'Maryam', 'g': 'A+', 's': 'Pass'},
        ]
        
        for std in students:
            g_val = ws.cell(row=std['row'], column=4).value
            s_val = ws.cell(row=std['row'], column=5).value
            
            correct_g = False
            if g_val:
                g_str = str(g_val).strip().upper().replace(" ", "")
                if g_str == std['g'].upper().replace(" ", ""):
                    correct_g = True
            
            correct_s = False
            if s_val:
                s_str = str(s_val).strip().lower()
                if s_str == std['s'].lower():
                    correct_s = True
            
            if correct_g and correct_s:
                score += 0.2
                details.append({'task': f'Student {std["name"]}', 'correct': True})
            else:
                errs = []
                if not correct_g: errs.append(f"Wrong Grade (Expected {std['g']})")
                if not correct_s: errs.append(f"Wrong Status (Expected {std['s']})")
                details.append({'task': f'Student {std["name"]}', 'correct': False, 'error': " & ".join(errs)})
                
    except Exception as e:
        details.append({'error': f'Grading error: {str(e)}'})
    return min(score, 2), details

def grade_complex(wb):
    score = 0
    details = []
    try:
        ws = wb['COMPLEX CHALLENGE']
        if ws['D20'].value and 'lap' in str(ws['D20'].value).lower(): score += 0.2; details.append({'q': 'Q21', 'correct': True})
        else: details.append({'q': 'Q21', 'correct': False})
        if ws['D21'].value and str(ws['D21'].value).strip() in ['5', '5.0']: score += 0.2; details.append({'q': 'Q22', 'correct': True})
        else: details.append({'q': 'Q22', 'correct': False})
        if ws['D22'].value and str(ws['D22'].value).strip() in ['25', '25.0']: score += 0.2; details.append({'q': 'Q23', 'correct': True})
        else: details.append({'q': 'Q23', 'correct': False})
        if ws['D23'].value and str(ws['D23'].value).strip() in ['4500', '4500.0']: score += 0.2; details.append({'q': 'Q24', 'correct': True})
        else: details.append({'q': 'Q24', 'correct': False})
        if ws['D24'].value and 'wireless' in str(ws['D24'].value).lower(): score += 0.2; details.append({'q': 'Q25', 'correct': True})
        else: details.append({'q': 'Q25', 'correct': False})
        for i in range(5):
            val = ws.cell(row=25+i, column=4).value
            if val: score += 0.2; details.append({'q': f'Q{26+i}', 'correct': True})
            else: details.append({'q': f'Q{26+i}', 'correct': False})
    except: pass
    return min(score, 2), details
