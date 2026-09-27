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
import re

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
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsm', dir=os.path.dirname(__file__))
        shutil.copy(template_path, temp_file.name)
        temp_file.close()
        wb = openpyxl.load_workbook(temp_file.name, keep_vba=True)
        os.unlink(temp_file.name)
    else:
        wb = openpyxl.Workbook()
    
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
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
    elif "Excel Skill 4" in assignment_title or "Advanced LOOKUP" in assignment_title:
        create_instructions_skill4(wb)
        create_lookup_function_exercises(wb)
        create_advanced_sumifs_exercises(wb)
        create_countifs_relationships_exercises(wb)
        create_integrated_lookup_challenge(wb)
    elif "Excel Skill 5" in assignment_title:
        create_instructions_skill5(wb) # Use new specialized instructions
        create_vlookup_sumif_countif_if_exercises(wb)
    else:
        create_instructions(wb)
        create_vlookup_exercises(wb)
        create_sumif_countif_exercises(wb)
        create_text_functions_exercises(wb)
        create_if_nested_exercises(wb)
        create_complex_challenge(wb)
    
    # Correct sheet visibility: hide 'Instructions', show exercise sheets
    if 'Instructions' in wb.sheetnames:
        wb['Instructions'].sheet_state = 'veryHidden'
    
    # Set the first visible sheet as active
    for sheetname in wb.sheetnames:
        if sheetname != 'Instructions':
            wb.active = wb[sheetname]
            break
    
    return wb

def create_instructions_skill5(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS: VLOOKUP, SUMIF, COUNTIF & IF"; ws['A1'].font = Font(size=18, bold=True, color="FFFFFF"); ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); ws.merge_cells('A1:F1'); ws.row_dimensions[1].height = 40
    ws['A2'] = "⚠️ IMPORTANT: YOU MUST ENABLE MACROS TO START"; ws['A2'].font = Font(size=14, bold=True, color="FF0000"); ws.merge_cells('A2:F2')
    ws['A3'] = "📋 OVERVIEW: Total Marks: 5. Perform all tasks in Yellow cells."; ws['A3'].font = Font(bold=True)
    ws['A5'] = "📝 FORMULA SOLUTIONS (Use these for reference):"; ws['A5'].font = Font(bold=True)
    ws['A6'] = "Q1 (VLOOKUP): =VLOOKUP(\"E050\", A4:E103, 5, 0)"
    ws['A7'] = "Q2 (SUMIF): =SUMIF(C4:C103, \"IT\", E4:E103)"
    ws['A8'] = "Q3 (COUNTIF): =COUNTIF(C4:C103, \"Karachi\")"
    ws['A9'] = "Q4 (IF): =IF(E108 > 45000, \"High\", \"Low\")"
    ws.column_dimensions['A'].width = 60

def create_lookup_function_exercises(wb):
    ws = wb.create_sheet("LOOKUP FUNCTION")
    ws['A1'] = "📝 Task 1: LOOKUP (Vector Form) - Use LOOKUP() function"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['G3'] = "REFERENCE TABLE (Sorted)"
    headers = ['Code', 'Product', 'Points']
    for col, h in enumerate(headers, 7): ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 3)
    ref_data = [[101, 'Apple', 10], [105, 'Banana', 20], [110, 'Cherry', 30], [120, 'Date', 40], [150, 'Elderberry', 50]]
    for i, row in enumerate(ref_data):
        for j, val in enumerate(row): ws.cell(row=5+i, column=7+j, value=val)
    ws['A4'] = "Q1. Find Product for code 110"; ws['B4'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A5'] = "Q2. Find Points for code 150"; ws['B5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A6'] = "Q3. Use LOOKUP to find 'Cherry' for code 110"; ws['B6'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['A'].width = 40; ws.column_dimensions['B'].width = 20

def create_advanced_sumifs_exercises(wb):
    ws = wb.create_sheet("ADVANCED SUMIFS")
    ws['A1'] = "📝 Task 2: SUMIFS with Multiple Criteria"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Date', 'Region', 'Category', 'Sales']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 4)
    data = [['2024-01-01', 'North', 'Electronics', 5000], ['2024-01-05', 'South', 'Furniture', 3000], ['2024-01-10', 'North', 'Furniture', 2000], ['2024-02-01', 'East', 'Electronics', 4000], ['2024-02-15', 'North', 'Electronics', 6000], ['2024-03-01', 'South', 'Electronics', 1500], ['2024-03-20', 'North', 'Furniture', 4500]]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    ws['F4'] = "Q4. Total Sales in 'North' for 'Electronics'"; ws['G4'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['F5'] = "Q5. Total Sales in 'South' for 'Furniture'"; ws['G5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['F'].width = 45; ws.column_dimensions['G'].width = 20

def create_countifs_relationships_exercises(wb):
    ws = wb.create_sheet("COUNTIFS & RELATIONSHIPS")
    ws['A1'] = "📝 Task 3: Relationships & COUNTIFS"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "Students Table"
    headers1 = ['SID', 'Name', 'DeptID']
    for col, h in enumerate(headers1, 1): ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 3)
    data1 = [[1, 'Ali', 'D1'], [2, 'Sara', 'D2'], [3, 'Zaman', 'D1'], [4, 'Bazan', 'D3'], [5, 'Taha', 'D1']]
    for i, row in enumerate(data1):
        for j, val in enumerate(row): ws.cell(row=5+i, column=1+j, value=val)
    ws['E3'] = "Departments Table"
    headers2 = ['DeptID', 'DeptName']
    for col, h in enumerate(headers2, 5): ws.cell(row=4, column=col, value=h)
    style_header(ws, 4, 2)
    data2 = [['D1', 'IT'], ['D2', 'HR'], ['D3', 'Sales']]
    for i, row in enumerate(data2):
        for j, val in enumerate(row): ws.cell(row=5+i, column=5+j, value=val)
    ws['A12'] = "Q6. Count students in 'D1' using COUNTIFS"; ws['B12'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A13'] = "Q7. Use LOOKUP to find DeptName for SID 4"; ws['B13'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['A'].width = 50; ws.column_dimensions['B'].width = 20

def create_integrated_lookup_challenge(wb):
    ws = wb.create_sheet("INTEGRATED CHALLENGE")
    ws['A1'] = "🏆 Final Challenge: The Master Report"; ws['A1'].font = Font(size=14, bold=True, color="C00000")
    ws['A3'] = "Instructions: Combine LOOKUP, SUMIFS, and COUNTIFS to answer these complex questions."
    ws['A5'] = "Q8. Total Sales for 'Electronics' in 'North' region after 2024-01-15"; ws['B5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A6'] = "Q9. How many departments have more than 2 students?"; ws['B6'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A7'] = "Q10. Use LOOKUP to find the 'Points' for 'Date'"; ws['B7'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['A'].width = 75; ws.column_dimensions['B'].width = 20

def create_instructions_dv(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS: DATA VALIDATION & NAME MANAGER"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1'); ws.row_dimensions[1].height = 40
    ws['A2'] = "⚠️ IMPORTANT: YOU MUST ENABLE MACROS TO START"
    ws['A2'].font = Font(size=14, bold=True, color="FF0000"); ws.merge_cells('A2:F2')
    ws['A4'] = "🚀 STEPS TO START:"; ws['A5'] = "1. Enable Content/Macros to see all sheets."; ws['A6'] = "2. Complete all tasks in YELLOW cells."; ws['A7'] = "3. AI will grade your submission based on correct validation and named ranges."
    ws['A9'] = "📋 ASSIGNMENT MODULES:"; ws['A10'] = "1. Named Manager (2.5 marks)"; ws['A11'] = "2. Dropdown Basic (2.5 marks)"; ws['A12'] = "3. Dropdown Advanced (2.5 marks)"; ws['A13'] = "4. Workbook Data Validation (2.5 marks)"
    ws.column_dimensions['A'].width = 55

def create_named_manager_exercises(wb):
    ws = wb.create_sheet("NAMED MANAGER")
    ws['A1'] = "📝 Task: Create Named Ranges"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "1. Select cells G4:G8 and name them 'ProductList'"
    ws['A4'] = "2. Select cells H4:H8 and name them 'PriceList'"
    ws['A5'] = "3. Select cells I4:I8 and name them 'CategoryList'"
    headers = ['Products', 'Prices', 'Categories']
    for col, h in enumerate(headers, 7): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 3)
    data = [['Laptop', 50000, 'Electronics'], ['Mouse', 1500, 'Accessories'], ['Keyboard', 3500, 'Accessories'], ['Monitor', 15000, 'Electronics'], ['USB', 1200, 'Storage']]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=7+j, value=val)

def create_dropdown_basic_exercises(wb):
    ws = wb.create_sheet("DROPDOWN BASIC")
    ws['A1'] = "📝 Task: Basic Dropdowns"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "1. In cell C5, create a dropdown using the list: Apple, Mango, Banana, Orange"
    ws['A4'] = "2. In cell C7, create a dropdown using the Named Range 'ProductList' from the previous sheet"
    ws['B5'] = "Select Fruit:"; ws['C5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B7'] = "Select Product:"; ws['C7'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 20; ws.column_dimensions['C'].width = 20

def create_dropdown_advanced_exercises(wb):
    ws = wb.create_sheet("DROPDOWN ADVANCED")
    ws['A1'] = "📝 Task: Dependent Dropdowns"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "1. Create Named Ranges for 'Electronics' (Laptop, Mobile) and 'Furniture' (Chair, Table)"
    ws['A4'] = "2. In cell C10, create a dropdown for Category (Electronics, Furniture)"
    ws['A5'] = "3. In cell C11, create a DEPENDENT dropdown that shows items based on C10"
    ws['G3'] = "Electronics"; ws['G4'] = "Laptop"; ws['G5'] = "Mobile"
    ws['H3'] = "Furniture"; ws['H4'] = "Chair"; ws['H5'] = "Table"
    ws['B10'] = "Category:"; ws['C10'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B11'] = "Item:"; ws['C11'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def create_workbook_structure_exercise(wb):
    ws = wb.create_sheet("DATA VALIDATION SKILL 2")
    ws['A1'] = "📝 Task: Whole Number & Date Validation"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "1. Apply validation to cell C5: Whole Number between 10 and 100"
    ws['A4'] = "2. Apply validation to cell C7: Date between 2024-01-01 and 2024-12-31"
    ws['A5'] = "3. Apply validation to cell C9: Text Length exactly 5 characters"
    ws['B5'] = "Enter Number (10-100):"; ws['C5'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B7'] = "Enter Date (2024):"; ws['C7'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['B9'] = "Enter Code (5 chars):"; ws['C9'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 30; ws.column_dimensions['C'].width = 30

def create_instructions_skill3(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS: DATA CLEANING & POWER QUERY"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1'); ws.row_dimensions[1].height = 40
    ws['A3'] = "🚀 STEPS TO START:"; ws['A4'] = "1. Enable Macros to see all task sheets."; ws['A5'] = "2. Task 1: Clean raw data using functions (TRIM, PROPER, etc.)."; ws['A6'] = "3. Task 2: Use Flash Fill or Text-to-Columns."; ws['A7'] = "4. Task 3: Handle Duplicates and Power Query transformation."
    ws['A9'] = "📋 ASSIGNMENT MODULES (Total 10 Marks):"; ws['A10'] = "1. Basic Data Cleaning (5 marks)"; ws['A11'] = "2. Power Query & Transformations (5 marks)"
    ws.column_dimensions['A'].width = 55

def create_data_cleaning_exercises(wb):
    ws = wb.create_sheet("DATA CLEANING")
    ws['A1'] = "📝 Task 1: Data Cleaning Functions (TRIM, PROPER, UPPER)"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Raw Data', 'Clean Data (Expected Formula)']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 2)
    raw_data = ["   muaaz asif   ", "PYTHON PROGRAMMING", "excel   skills", "john doe  ", "   KARAchi-PAKistan"]
    for i, data in enumerate(raw_data):
        ws.cell(row=4+i, column=1, value=data)
        ws.cell(row=4+i, column=2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws['A11'] = "📝 Task 2: Flash Fill / Text-to-Columns"
    ws['A12'] = "Separate 'First Name' and 'Last Name' from the Full Name column."
    ws['A14'] = "Full Name"; ws['B14'] = "First Name"; ws['C14'] = "Last Name"
    style_header(ws, 14, 3)
    names = ["Ali Khan", "Sara Ahmed", "Bilal Sheikh"]
    for i, name in enumerate(names):
        ws.cell(row=15+i, column=1, value=name)
        ws.cell(row=15+i, column=2).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws.cell(row=15+i, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['A'].width = 30; ws.column_dimensions['B'].width = 30; ws.column_dimensions['C'].width = 30

def create_power_query_exercises(wb):
    ws = wb.create_sheet("POWER QUERY")
    ws['A1'] = "📝 Task 3: Data Transformation & Duplicates"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    ws['A3'] = "1. Identify and remove duplicate rows from the table below (F4:H10)."
    ws['A4'] = "2. Ensure all text is in UPPERCASE in the 'Product' column."
    data = [['ID', 'Product', 'Sales'], [101, 'Laptop', 500], [102, 'Mouse', 50], [101, 'Laptop', 500], [103, 'Keyboard', 80], [102, 'Mouse', 50], [104, 'Monitor', 300]]
    for i, row in enumerate(data):
        for j, val in enumerate(row):
            ws.cell(row=4+i, column=6+j, value=val)
            if i > 0: ws.cell(row=4+i, column=6+j).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    style_header(ws, 4, 3); ws.column_dimensions['F'].width = 10; ws.column_dimensions['G'].width = 20; ws.column_dimensions['H'].width = 15

def create_instructions(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws['A1'] = "📊 EXCEL SKILLS ASSIGNMENT - Complete Workbook"; ws['A1'].font = Font(size=18, bold=True, color="FFFFFF"); ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid"); ws.merge_cells('A1:F1'); ws.row_dimensions[1].height = 40
    ws['A2'] = "⚠️ IMPORTANT: YOU MUST ENABLE MACROS TO START"; ws['A2'].font = Font(size=14, bold=True, color="FF0000"); ws.merge_cells('A2:F2')
    ws['A3'] = "🚀 STEPS TO START:"; ws['A3'].font = Font(bold=True); ws['A4'] = "1. Enable Content/Macros to see all task sheets."; ws['A5'] = "2. Write formulas in YELLOW cells."; ws['A6'] = "3. AI will grade and provide instant feedback."
    ws['A8'] = "🚨 ANTI-CHEATING: IF YOU OPEN ANY OTHER WINDOW, YOUR MARKS WILL BE ZERO!"; ws['A8'].font = Font(size=11, bold=True, color="C00000"); ws.merge_cells('A8:F8')
    content = [("", None), ("📋 OVERVIEW:", None), ("• Total Marks: 10", None), ("• Time Limit: 2 hours", None), ("", None), ("📝 EXERCISE SHEETS:", None), ("1. VLOOKUP (2 marks)", None), ("2. SUMIF & COUNTIF (2 marks)", None), ("3. LEFT, RIGHT, MID (2 marks)", None), ("4. IF & NESTED IF (2 marks)", None), ("5. COMPLEX CHALLENGE (2 marks)", None)]
    for i, (text, _) in enumerate(content, 5): ws.cell(row=i, column=1, value=text)
    ws.column_dimensions['A'].width = 55

def create_vlookup_exercises(wb):
    ws = wb.create_sheet("VLOOKUP")
    ws.merge_cells('A1:E1'); ws['A1'] = "📊 Employee Database"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Emp ID', 'Name', 'Department', 'City', 'Salary']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 5)
    data = [['E001', 'Ahmed Ali', 'IT', 'Karachi', 45000], ['E002', 'Sara Khan', 'HR', 'Lahore', 42000], ['E003', 'Omar Sheikh', 'Finance', 'Islamabad', 50000], ['E004', 'Fatima Noor', 'IT', 'Karachi', 48000], ['E005', 'Bilal Ahmed', 'Sales', 'Peshawar', 38000], ['E006', 'Ayesha Malik', 'HR', 'Lahore', 41000], ['E007', 'Hassan Raza', 'Finance', 'Islamabad', 52000], ['E008', 'Zainab Hussain', 'Sales', 'Quetta', 36000], ['E009', 'Ali Raza', 'IT', 'Karachi', 46000], ['E010', 'Maryam Fatima', 'HR', 'Lahore', 43000]]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    q_row = 16; ws.merge_cells(f'A{q_row}:F{q_row}'); ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use VLOOKUP formula").font = Font(size=12, bold=True, color="C00000")
    questions = [['Q1', 'Find Department of Employee E003'], ['Q2', 'Find Salary of Sara Khan'], ['Q3', 'Find City of Employee E007'], ['Q4', 'Find Name of Employee with ID E010']]
    for i, q in enumerate(questions):
        r = q_row + 3 + i; ws.cell(row=r, column=1, value=q[0]); ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 45

def create_sumif_countif_exercises(wb):
    ws = wb.create_sheet("SUMIF & COUNTIF")
    ws.merge_cells('A1:F1'); ws['A1'] = "📊 Sales Data - Q1 2024"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Date', 'Salesperson', 'Product', 'Category', 'Quantity', 'Amount']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 6)
    data = [['2024-01-05', 'Ali', 'Laptop', 'Electronics', 2, 120000], ['2024-01-10', 'Sara', 'Mouse', 'Accessories', 10, 15000], ['2024-01-15', 'Ali', 'Keyboard', 'Accessories', 5, 25000], ['2024-02-01', 'Omar', 'Laptop', 'Electronics', 3, 180000], ['2024-02-10', 'Sara', 'Monitor', 'Electronics', 4, 80000], ['2024-02-15', 'Ali', 'Mouse', 'Accessories', 8, 12000], ['2024-03-01', 'Omar', 'Keyboard', 'Accessories', 6, 30000], ['2024-03-10', 'Fatima', 'Laptop', 'Electronics', 2, 120000], ['2024-03-15', 'Sara', 'Laptop', 'Electronics', 1, 60000], ['2024-03-20', 'Fatima', 'Monitor', 'Electronics', 3, 60000]]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    q_row = 17; ws.merge_cells(f'A{q_row}:F{q_row}'); ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use SUMIF and COUNTIF").font = Font(size=12, bold=True, color="C00000")
    questions = [['Q5', 'Total sales by Ali (SUMIF)'], ['Q6', 'Count of Laptop sales (COUNTIF)'], ['Q7', 'Total quantity sold by Sara (SUMIF)'], ['Q8', 'Count of Electronics sold (COUNTIF)'], ['Q9', 'Total amount of Accessories (SUMIF)'], ['Q10', 'Count of Fatima sales (COUNTIF)']]
    for i, q in enumerate(questions):
        r = q_row + 3 + i; ws.cell(row=r, column=1, value=q[0]); ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 45

def create_text_functions_exercises(wb):
    ws = wb.create_sheet("LEFT RIGHT MID")
    ws.merge_cells('A1:D1'); ws['A1'] = "📊 Student Data"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Full Name', 'Phone Number', 'Email', 'Student Code']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 4)
    data = [['Ahmed Ali Khan', '0300-1234567', 'ahmed.ali@gmail.com', 'STU-2024-001'], ['Sara Fatima', '0321-7654321', 'sara.f@hotmail.com', 'STU-2024-002'], ['Muhammad Omar', '0333-9876543', 'omar.pk@yahoo.com', 'STU-2024-003'], ['Fatima Noor', '0345-1112233', 'fatima.noor@outlook.com', 'STU-2024-004'], ['Bilal Hussain', '0301-4445566', 'bilal.h@gmail.com', 'STU-2024-005']]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    q_row = 12; ws.merge_cells(f'A{q_row}:E{q_row}'); ws.cell(row=q_row, column=1, value="📝 EXERCISES - LEFT, RIGHT, MID, LEN, FIND").font = Font(size=12, bold=True, color="C00000")
    questions = [['Q11', 'Extract first 3 letters from A4'], ['Q12', 'Extract last 7 digits from B4'], ['Q13', 'Extract username from C4'], ['Q14', 'Extract year from D4'], ['Q15', 'Extract domain from C5'], ['Q16', 'Extract middle name from A5']]
    for i, q in enumerate(questions):
        r = q_row + 3 + i; ws.cell(row=r, column=1, value=q[0]); ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 50

def create_if_nested_exercises(wb):
    ws = wb.create_sheet("IF & NESTED IF")
    ws.merge_cells('A1:E1'); ws['A1'] = "📊 Student Marks Data"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Student ID', 'Name', 'Marks', 'Grade (IF)', 'Status (IF)']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 5)
    data = [['S001', 'Ahmed', 92, '', ''], ['S002', 'Sara', 85, '', ''], ['S003', 'Omar', 76, '', ''], ['S004', 'Fatima', 65, '', ''], ['S005', 'Bilal', 58, '', ''], ['S006', 'Ayesha', 45, '', ''], ['S007', 'Hassan', 35, '', ''], ['S008', 'Zainab', 88, '', ''], ['S009', 'Ali', 72, '', ''], ['S010', 'Maryam', 95, '', '']]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    grading = ["90-100 = A+", "80-89 = A", "70-79 = B", "60-69 = C", "50-59 = D", "Below 50 = F", "", ">= 50 = Pass", "< 50 = Fail"]
    for i, text in enumerate(grading, 4): ws.cell(row=3+i, column=7, value=text)
    ws.column_dimensions['B'].width = 20

def create_complex_challenge(wb):
    ws = wb.create_sheet("COMPLEX CHALLENGE")
    ws.merge_cells('A1:F1'); ws['A1'] = "🏆 CHALLENGE - Complete Product Analysis"; ws['A1'].font = Font(size=14, bold=True, color="C00000")
    headers = ['Product Code', 'Product Name', 'Category', 'Price', 'Stock', 'Status']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 6)
    data = [['PRD-LAP-001', 'Laptop Pro 15', 'Electronics', 85000, 12, ''], ['PRD-MOU-002', 'Wireless Mouse', 'Accessories', 1500, 50, ''], ['PRD-KEY-003', 'Mech Keyboard', 'Accessories', 4500, 25, ''], ['PRD-MON-004', 'Monitor 27 inch', 'Electronics', 35000, 8, ''], ['PRD-USB-005', 'USB Hub 7-in-1', 'Accessories', 2500, 40, ''], ['PRD-HDD-006', 'External HDD 1TB', 'Storage', 8000, 15, ''], ['PRD-SSD-007', 'SSD 500GB', 'Storage', 6500, 20, ''], ['PRD-WEB-008', 'Webcam HD', 'Accessories', 3500, 30, ''], ['PRD-TAB-009', 'Tablet 10 inch', 'Electronics', 45000, 5, ''], ['PRD-SPK-010', 'Bluetooth Speaker', 'Accessories', 5500, 35, '']]
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    q_row = 17; ws.merge_cells(f'A{q_row}:F{q_row}'); ws.cell(row=q_row, column=1, value="📝 CHALLENGE - Combine all functions!").font = Font(size=12, bold=True, color="C00000")
    questions = [['Q21', 'Extract category code from A4'], ['Q22', 'Count products in "Accessories"'], ['Q23', 'Total stock of Electronics'], ['Q24', 'Find price of PRD-KEY-003'], ['Q25', 'Extract first word from B5'], ['Q26', 'Status: IF stock > 20 then "In Stock" else "Low"'], ['Q27', 'Total value of all Accessories'], ['Q28', 'Extract number from A5'], ['Q29', 'Count products with price > 10000'], ['Q30', 'Extract domain from admin@company.com']]
    for i, q in enumerate(questions):
        r = q_row + 3 + i; ws.cell(row=r, column=1, value=q[0]); ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 55

def create_vlookup_sumif_countif_if_exercises(wb):
    ws = wb.create_sheet("EXCEL SKILL 5")
    ws.merge_cells('A1:F1'); ws['A1'] = "📊 Student Data (100 Students)"; ws['A1'].font = Font(size=14, bold=True, color="1F4E79")
    headers = ['Emp ID', 'Name', 'Department', 'City', 'Salary']
    for col, h in enumerate(headers, 1): ws.cell(row=3, column=col, value=h)
    style_header(ws, 3, 5)
    
    # Generate 100 students
    data = []
    departments = ['IT', 'HR', 'Finance', 'Sales', 'Marketing']
    cities = ['Karachi', 'Lahore', 'Islamabad', 'Peshawar', 'Quetta']
    import random
    for i in range(1, 101):
        data.append([f'E{i:03d}', f'Student {i}', random.choice(departments), random.choice(cities), random.randint(30000, 60000)])
        
    for i, row in enumerate(data):
        for j, val in enumerate(row): ws.cell(row=4+i, column=1+j, value=val)
    
    # Exercises
    q_row = 106
    ws.merge_cells(f'A{q_row}:F{q_row}'); ws.cell(row=q_row, column=1, value="📝 EXERCISES (Use formulas in Yellow cells)").font = Font(size=12, bold=True, color="C00000")
    questions = [
        ['Q1', 'Find Salary of Employee E050 (VLOOKUP)', 'B108'],
        ['Q2', 'Total Salary of IT Department (SUMIF)', 'B109'],
        ['Q3', 'Count of students from Karachi (COUNTIF)', 'B110'],
        ['Q4', 'If Salary > 45000 then "High" else "Low" (IF)', 'B111']
    ]
    for i, q in enumerate(questions):
        r = q_row + 2 + i; ws.cell(row=r, column=1, value=q[0]); ws.cell(row=r, column=2, value=q[1])
        ws.cell(row=r, column=3).fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws.column_dimensions['B'].width = 50

# AUTO-GRADER

def grade_excel_submission(file_path, assignment_title=""):
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False)
    except Exception as e: return {'error': f'Cannot open file: {str(e)}', 'score': 0}
    cheating_detected = False; macros_disabled = False
    if 'Instructions' in wb.sheetnames:
        ws = wb['Instructions']; macro_flag = ws.cell(row=99, column=26).value; cheat_flag = ws.cell(row=100, column=26).value
        if not macro_flag or str(macro_flag).upper() != 'MACROS_OK': macros_disabled = True
        if cheat_flag and 'CHEAT' in str(cheat_flag).upper(): cheating_detected = True
    if cheating_detected: return {'score': 0, 'max': 10, 'percentage': 0, 'cheating_detected': True, 'details': {'error': 'CHEATING DETECTED'}}
    total_score = 0; details = {}
    if "Data Validation" in assignment_title and "Manager" in assignment_title:
        nm_score, nm_detail = grade_named_manager(wb); db_score, db_detail = grade_dropdown_basic(wb); da_score, da_detail = grade_dropdown_advanced(wb); wv_score, wv_detail = grade_workbook_validation(wb)
        total_score = nm_score + db_score + da_score + wv_score
        details['Named Manager'] = {'score': nm_score, 'max': 2.5, 'details': nm_detail}
        details['Dropdown Basic'] = {'score': db_score, 'max': 2.5, 'details': db_detail}
        details['Dropdown Advanced'] = {'score': da_score, 'max': 2.5, 'details': da_detail}
        details['Workbook Validation'] = {'score': wv_score, 'max': 2.5, 'details': wv_detail}
    elif "Data Cleaning" in assignment_title and "Power Query" in assignment_title:
        wb_vals = openpyxl.load_workbook(file_path, data_only=True)
        dc_score, dc_detail = grade_data_cleaning(wb_vals); pq_score, pq_detail = grade_power_query(wb_vals)
        total_score = dc_score + pq_score
        details['Data Cleaning'] = {'score': dc_score, 'max': 5, 'details': dc_detail}
        details['Power Query Basics'] = {'score': pq_score, 'max': 5, 'details': pq_detail}
    elif "Excel Skill 4" in assignment_title or "Advanced LOOKUP" in assignment_title:
        wb_vals = openpyxl.load_workbook(file_path, data_only=True)
        l_score, l_detail = grade_lookup_function(wb_vals); s_score, s_detail = grade_advanced_sumifs(wb_vals); r_score, r_detail = grade_countifs_relationships(wb_vals); i_score, i_detail = grade_integrated_lookup(wb_vals)
        total_score = l_score + s_score + r_score + i_score
        details['LOOKUP Function'] = {'score': l_score, 'max': 2.5, 'details': l_detail}
        details['Advanced SUMIFS'] = {'score': s_score, 'max': 2.5, 'details': s_detail}
        details['COUNTIFS & Relationships'] = {'score': r_score, 'max': 2.5, 'details': r_detail}
        details['Integrated Challenge'] = {'score': i_score, 'max': 2.5, 'details': i_detail}
    elif "Excel Skill 5" in assignment_title:
        v_score, v_detail = grade_vlookup_sumif_countif_if(wb)
        total_score = v_score
        details['VLOOKUP/SUMIF/COUNTIF/IF'] = {'score': v_score, 'max': 5, 'details': v_detail}
    else:
        score_s1, details_s1 = grade_excel_skill_1(wb)
        total_score = score_s1
        details['Excel Skill 1 (Q1-Q20)'] = {'score': score_s1, 'max': 10, 'details': details_s1}
    return {'score': round(min(total_score, 10), 2), 'max': 10, 'percentage': round((min(total_score, 10) / 10) * 100, 1), 'cheating_detected': False, 'details': details}

QUESTION_MARK = 0.5

def normalize_formula(value):
    if value is None:
        return ""
    s = str(value).strip().upper()
    s = re.sub(r"\s+", "", s)
    if s.startswith("="):
        s = s[1:]
    s = s.replace("$", "")
    return s

def is_formula(value):
    return isinstance(value, str) and value.strip().startswith("=")

def get_function(formula):
    f = normalize_formula(formula)
    m = re.match(r"^([A-Z][A-Z0-9_.]*)\(", f)
    return m.group(1) if m else ""

ANSWER_CELLS_S1 = {
    "Q1": ("VLOOKUP", "C19"),
    "Q2": ("VLOOKUP", "C20"),
    "Q3": ("VLOOKUP", "C21"),
    "Q4": ("VLOOKUP", "C22"),

    "Q5": ("SUMIF & COUNTIF", "C20"),
    "Q6": ("SUMIF & COUNTIF", "C21"),
    "Q7": ("SUMIF & COUNTIF", "C22"),
    "Q8": ("SUMIF & COUNTIF", "C23"),
    "Q9": ("SUMIF & COUNTIF", "C24"),
    "Q10": ("SUMIF & COUNTIF", "C25"),

    "Q11": ("LEFT RIGHT MID", "C15"),
    "Q12": ("LEFT RIGHT MID", "C16"),
    "Q13": ("LEFT RIGHT MID", "C17"),
    "Q14": ("LEFT RIGHT MID", "C18"),
    "Q15": ("LEFT RIGHT MID", "C19"),
    "Q16": ("LEFT RIGHT MID", "C20"),

    "Q17": ("Complex Challenge", "C17"),
    "Q18": ("Complex Challenge", "C18"),
    "Q19": ("Complex Challenge", "C19"),
    "Q20": ("Complex Challenge", "C20"),
}

def normalize_range_list(ranges):
    return {
        normalize_formula(r)
        for r in ranges
    }

def check_vlookup_s1(formula, lookup_values, return_col, table_ranges):
    result = {"marks": 0, "status": "WRONG", "issues": [], "checks": {}}
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result
    if get_function(formula) != "VLOOKUP":
        result["issues"].append("VLOOKUP is required.")
        return result
    args = split_formula_args(formula)
    if len(args) < 4:
        result["issues"].append("VLOOKUP needs 4 arguments.")
        return result
    lookup = normalize_formula(args[0])
    table = normalize_formula(args[1])
    col = normalize_formula(args[2])
    match = normalize_formula(args[3])
    result["checks"]["function"] = 1
    
    possible_lookup_values = set()
    if isinstance(lookup_values, (list, tuple, set)):
        for value in lookup_values:
            norm_val = normalize_formula(value)
            possible_lookup_values.add(norm_val)
            possible_lookup_values.add(f'"{norm_val}"')
    else:
        norm_val = normalize_formula(lookup_values)
        possible_lookup_values.add(norm_val)
        possible_lookup_values.add(f'"{norm_val}"')

    possible_table_ranges = normalize_range_list(table_ranges if isinstance(table_ranges, list) else [table_ranges])

    result["checks"]["lookup_value"] = int(lookup in possible_lookup_values)
    result["checks"]["table_range"] = int(table in possible_table_ranges)
    result["checks"]["return_column"] = int(col == str(return_col))
    result["checks"]["exact_match"] = int(match in {"0", "FALSE"})
    for name, value in result["checks"].items():
        if value == 0:
            result["issues"].append(name + " is incorrect.")
    if all(v == 1 for v in result["checks"].values()):
        result["marks"] = QUESTION_MARK
        result["status"] = "CORRECT"
    return result

def check_sumif_s1(formula, criteria_ranges, criteria, sum_ranges):
    result = {"marks": 0, "status": "WRONG", "issues": [], "checks": {}}
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result
    if get_function(formula) != "SUMIF":
        result["issues"].append("SUMIF is required.")
        return result
    args = split_formula_args(formula)
    if len(args) != 3:
        result["issues"].append("SUMIF must have 3 arguments.")
        return result
    criteria_range = normalize_formula(args[0])
    actual_criteria = normalize_formula(args[1])
    sum_range = normalize_formula(args[2])

    possible_criteria_ranges = normalize_range_list(criteria_ranges if isinstance(criteria_ranges, list) else [criteria_ranges])
    possible_sum_ranges = normalize_range_list(sum_ranges if isinstance(sum_ranges, list) else [sum_ranges])
    possible_criteria = {normalize_formula(criteria), f'"{normalize_formula(criteria)}"'}

    result["checks"]["function"] = 1
    result["checks"]["criteria_range"] = int(criteria_range in possible_criteria_ranges)
    result["checks"]["criteria"] = int(actual_criteria in possible_criteria)
    result["checks"]["sum_range"] = int(sum_range in possible_sum_ranges)
    for name, value in result["checks"].items():
        if value == 0:
            result["issues"].append(name + " is incorrect.")
    if all(v == 1 for v in result["checks"].values()):
        result["marks"] = QUESTION_MARK
        result["status"] = "CORRECT"
    return result

def check_countif_s1(formula, criteria_ranges, criteria):
    result = {"marks": 0, "status": "WRONG", "issues": [], "checks": {}}
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result
    if get_function(formula) != "COUNTIF":
        result["issues"].append("COUNTIF is required.")
        return result
    args = split_formula_args(formula)
    if len(args) != 2:
        result["issues"].append("COUNTIF must have 2 arguments.")
        return result
    criteria_range = normalize_formula(args[0])
    actual_criteria = normalize_formula(args[1])

    possible_criteria_ranges = normalize_range_list(criteria_ranges if isinstance(criteria_ranges, list) else [criteria_ranges])
    possible_criteria = {normalize_formula(criteria), f'"{normalize_formula(criteria)}"'}

    result["checks"]["function"] = 1
    result["checks"]["criteria_range"] = int(criteria_range in possible_criteria_ranges)
    result["checks"]["criteria"] = int(actual_criteria in possible_criteria)
    for name, value in result["checks"].items():
        if value == 0:
            result["issues"].append(name + " is incorrect.")
    if all(v == 1 for v in result["checks"].values()):
        result["marks"] = QUESTION_MARK
        result["status"] = "CORRECT"
    return result

def check_text_formula_s1(formula, accepted_patterns, expected_text):
    result = {"marks": 0, "status": "WRONG", "issues": [], "checks": {}}
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result
    f = normalize_formula(formula)
    result["checks"]["formula_structure"] = int(any(re.fullmatch(p, f) for p in accepted_patterns))
    if result["checks"]["formula_structure"]:
        result["marks"] = QUESTION_MARK
        result["status"] = "CORRECT"
    else:
        result["issues"].append(f"Formula does not match accepted solution for: {expected_text}")
    return result

def check_q1_s1(f): return check_vlookup_s1(f, ["E003"], 3, ["A4:C13", "A3:C13"])
def check_q2_s1(f): return check_vlookup_s1(f, ["Sara Khan", "B20"], 4, ["B4:E13", "B3:E13"])
def check_q3_s1(f): return check_vlookup_s1(f, ["E007"], 4, ["A4:D13", "A3:D13"])
def check_q4_s1(f): return check_vlookup_s1(f, ["E010"], 2, ["A4:B13", "A3:B13"])

def check_q5_s1(f): return check_sumif_s1(f, ["B4:B13", "B3:B13"], "ALI", ["F4:F13", "F3:F13"])
def check_q6_s1(f): return check_countif_s1(f, ["C4:C13", "C3:C13"], "LAPTOP")
def check_q7_s1(f): return check_sumif_s1(f, ["B4:B13", "B3:B13"], "SARA", ["E4:E13", "E3:E13"])
def check_q8_s1(f): return check_countif_s1(f, ["D4:D13", "D3:D13"], "ELECTRONICS")
def check_q9_s1(f): return check_sumif_s1(f, ["D4:D13", "D3:D13"], "ACCESSORIES", ["F4:F13", "F3:F13"])
def check_q10_s1(f): return check_countif_s1(f, ["B4:B13", "B3:B13"], "FATIMA")

def check_q11_s1(f): return check_text_formula_s1(f, [r'LEFT\(A4,3\)'], "Ahm")
def check_q12_s1(f): return check_text_formula_s1(f, [r'RIGHT\(B4,7\)'], "1234567")
def check_q13_s1(f): return check_text_formula_s1(f, [r'LEFT\(C4,FIND\("@",C4\)-1\)'], "ahmed.ali")
def check_q14_s1(f): return check_text_formula_s1(f, [r'MID\(D4,5,4\)'], "2024")
def check_q15_s1(f): return check_text_formula_s1(f, [r'RIGHT\(C5,LEN\(C5\)-FIND\("@",C5\)\)', r'MID\(C5,FIND\("@",C5\)+1,LEN\(C5\)\)'], "hotmail.com")
def check_q16_s1(f): return check_text_formula_s1(f, [r'RIGHT\(A5,LEN\(A5\)-FIND\(" ",A5\)\)', r'MID\(A5,FIND\(" ",A5\)+1,LEN\(A5\)\)'], "Fatima")

def check_q17_s1(f): return check_text_formula_s1(f, [r'D4\*E4', r'E4\*D4'], "340000")
def check_q18_s1(f): return check_text_formula_s1(f, [r'RIGHT\(A4,3\)', r'MID\(A4,9,3\)'], "LAP")
def check_q19_s1(f): return check_countif_s1(f, ["C4:C11", "C3:C11"], "ACCESSORIES")
def check_q20_s1(f): return check_sumif_s1(f, ["C4:C11", "C3:C11"], "ELECTRONICS", ["F4:F11", "F3:F11"])

CHECKERS_S1 = {
    "Q1": check_q1_s1, "Q2": check_q2_s1, "Q3": check_q3_s1, "Q4": check_q4_s1,
    "Q5": check_q5_s1, "Q6": check_q6_s1, "Q7": check_q7_s1, "Q8": check_q8_s1, "Q9": check_q9_s1, "Q10": check_q10_s1,
    "Q11": check_q11_s1, "Q12": check_q12_s1, "Q13": check_q13_s1, "Q14": check_q14_s1, "Q15": check_q15_s1, "Q16": check_q16_s1,
    "Q17": check_q17_s1, "Q18": check_q18_s1, "Q19": check_q19_s1, "Q20": check_q20_s1,
}

def grade_excel_skill_1(wb):
    total_score = 0; details = []
    for q_id, (sheet_name, cell_ref) in ANSWER_CELLS_S1.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            formula = ws[cell_ref].value
            try:
                res = CHECKERS_S1[q_id](formula)
                marks = res.get('marks', 0)
                status = res.get('status', 'WRONG')
                issues = res.get('issues', [])
                total_score += marks
                details.append({
                    'q': q_id,
                    'task': f"Question {q_id} ({sheet_name} {cell_ref})",
                    'correct': status == 'CORRECT',
                    'error': " | ".join(issues) if issues else "Formula + result correct (0.5/0.5)"
                })
            except Exception as e:
                details.append({
                    'q': q_id,
                    'task': f"Question {q_id} ({sheet_name} {cell_ref})",
                    'correct': False,
                    'error': str(e)
                })
        else:
            details.append({
                'q': q_id,
                'task': f"Question {q_id} ({sheet_name})",
                'correct': False,
                'error': f"Sheet '{sheet_name}' not found."
            })
    return round(total_score, 2), details

# GRADING HELPERS FOR SKILL 4

def grade_lookup_function(wb):
    score = 0; details = []
    try:
        ws = wb['LOOKUP FUNCTION']
        v1 = ws['B4'].value
        if v1 and 'cherry' in str(v1).lower():
            score += 0.8; details.append({'q': 'Q1', 'task': 'Find Product for code 110', 'correct': True})
        else:
            details.append({'q': 'Q1', 'task': 'Find Product for code 110', 'correct': False, 'error': f"Expected 'Cherry', Got '{v1}' (Cell B4)"})

        v2 = ws['B5'].value
        if v2 and str(v2).strip() in ['50', '50.0']:
            score += 0.8; details.append({'q': 'Q2', 'task': 'Find Points for code 150', 'correct': True})
        else:
            details.append({'q': 'Q2', 'task': 'Find Points for code 150', 'correct': False, 'error': f"Expected '50', Got '{v2}' (Cell B5)"})

        v3 = ws['B6'].value
        if v3 and 'cherry' in str(v3).lower():
            score += 0.9; details.append({'q': 'Q3', 'task': 'LOOKUP function for code 110', 'correct': True})
        else:
            details.append({'q': 'Q3', 'task': 'LOOKUP function for code 110', 'correct': False, 'error': f"Expected 'Cherry', Got '{v3}' (Cell B6)"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2.5), details

def grade_advanced_sumifs(wb):
    score = 0; details = []
    try:
        ws = wb['ADVANCED SUMIFS']
        v4 = ws['G4'].value
        if v4 and str(v4).strip() in ['11000', '11000.0']:
            score += 1.25; details.append({'q': 'Q4', 'task': "Total Sales in 'North' for 'Electronics'", 'correct': True})
        else:
            details.append({'q': 'Q4', 'task': "Total Sales in 'North' for 'Electronics'", 'correct': False, 'error': f"Expected '11000', Got '{v4}' (Cell G4)"})

        v5 = ws['G5'].value
        if v5 and str(v5).strip() in ['3000', '3000.0']:
            score += 1.25; details.append({'q': 'Q5', 'task': "Total Sales in 'South' for 'Furniture'", 'correct': True})
        else:
            details.append({'q': 'Q5', 'task': "Total Sales in 'South' for 'Furniture'", 'correct': False, 'error': f"Expected '3000', Got '{v5}' (Cell G5)"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2.5), details

def grade_countifs_relationships(wb):
    score = 0; details = []
    try:
        ws = wb['COUNTIFS & RELATIONSHIPS']
        v6 = ws['B12'].value
        if v6 and str(v6).strip() in ['3', '3.0']:
            score += 1.25; details.append({'q': 'Q6', 'task': "Count students in 'D1' using COUNTIFS", 'correct': True})
        else:
            details.append({'q': 'Q6', 'task': "Count students in 'D1' using COUNTIFS", 'correct': False, 'error': f"Expected '3', Got '{v6}' (Cell B12)"})

        v7 = ws['B13'].value
        if v7 and 'sales' in str(v7).lower():
            score += 1.25; details.append({'q': 'Q7', 'task': "LOOKUP to find DeptName for SID 4", 'correct': True})
        else:
            details.append({'q': 'Q7', 'task': "LOOKUP to find DeptName for SID 4", 'correct': False, 'error': f"Expected 'Sales', Got '{v7}' (Cell B13)"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2.5), details

def grade_integrated_lookup(wb):
    score = 0; details = []
    try:
        ws = wb['INTEGRATED CHALLENGE']
        v8 = ws['B5'].value
        if v8 and str(v8).strip() in ['6000', '6000.0']:
            score += 1.0; details.append({'q': 'Q8', 'task': "Total Sales 'Electronics' in 'North' after date", 'correct': True})
        else:
            details.append({'q': 'Q8', 'task': "Total Sales 'Electronics' in 'North' after date", 'correct': False, 'error': f"Expected '6000', Got '{v8}' (Cell B5)"})

        v9 = ws['B6'].value
        if v9 and str(v9).strip() in ['1', '1.0']:
            score += 0.75; details.append({'q': 'Q9', 'task': "Departments with more than 2 students", 'correct': True})
        else:
            details.append({'q': 'Q9', 'task': "Departments with more than 2 students", 'correct': False, 'error': f"Expected '1', Got '{v9}' (Cell B6)"})

        v10 = ws['B7'].value
        if v10 and str(v10).strip() in ['40', '40.0']:
            score += 0.75; details.append({'q': 'Q10', 'task': "LOOKUP points for date", 'correct': True})
        else:
            details.append({'q': 'Q10', 'task': "LOOKUP points for date", 'correct': False, 'error': f"Expected '40', Got '{v10}' (Cell B7)"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2.5), details

# OTHER GRADING HELPERS

def grade_data_cleaning(wb):
    score = 0; details = []
    try:
        ws = wb['DATA CLEANING']; expected = ["Muaaz Asif", "Python Programming", "Excel Skills", "John Doe", "Karachi-Pakistan"]
        for i, exp in enumerate(expected):
            val = ws.cell(row=4+i, column=2).value
            if val and str(val).strip().lower() == exp.lower(): score += 0.6; details.append({'task': f'Clean: {exp}', 'correct': True})
            else: details.append({'task': f'Clean: {exp}', 'correct': False})
        names = [("Ali", "Khan"), ("Sara", "Ahmed"), ("Bilal", "Sheikh")]
        for i, (f, l) in enumerate(names):
            fv = ws.cell(row=15+i, column=2).value; lv = ws.cell(row=15+i, column=3).value
            if fv and str(fv).strip().lower() == f.lower() and lv and str(lv).strip().lower() == l.lower(): score += 0.66; details.append({'task': f'Split: {f} {l}', 'correct': True})
            else: details.append({'task': f'Split: {f} {l}', 'correct': False})
    except: pass
    return min(score, 5), details

def grade_power_query(wb):
    score = 0; details = []
    try:
        ws = wb['POWER QUERY']; products = []
        for r in range(5, 11):
            p = ws.cell(row=r, column=7).value
            if p: products.append(str(p).strip().upper())
        if len(products) == 5 and "LAPTOP" in products and "MOUSE" in products: score += 2.5; details.append({'task': 'Duplicates Removed', 'correct': True})
        else: details.append({'task': 'Duplicates Removed', 'correct': False})
        if all(p.isupper() for p in products if p): score += 2.5; details.append({'task': 'Uppercase', 'correct': True})
        else: details.append({'task': 'Uppercase', 'correct': False})
    except: pass
    return min(score, 5), details

import excel_auto_checker_skill2

def grade_named_manager(wb):
    score = 0; details = []
    try:
        for name, (sheet_name, cell_range) in excel_auto_checker_skill2.NAMED_MANAGER_CHECKS.items():
            check = excel_auto_checker_skill2.check_named_range(wb, name, sheet_name, cell_range)
            correct = (check["status"] == "CORRECT")
            if correct:
                score += excel_auto_checker_skill2.NAMED_MANAGER_MARKS / len(excel_auto_checker_skill2.NAMED_MANAGER_CHECKS)
            details.append({
                'task': name,
                'correct': correct,
                'error': " | ".join(check["issues"]) if not correct else ""
            })
    except Exception as e:
        details.append({'error': str(e)})
    return round(min(score, 2.5), 4), details

def grade_dropdown_basic(wb):
    score = 0; details = []
    try:
        if "DROPDOWN BASIC" not in wb.sheetnames:
            details.append({'task': 'C5', 'correct': False, 'error': "Sheet 'DROPDOWN BASIC' not found."})
            details.append({'task': 'C7', 'correct': False, 'error': "Sheet 'DROPDOWN BASIC' not found."})
        else:
            ws = wb["DROPDOWN BASIC"]
            c5_check = excel_auto_checker_skill2.check_basic_c5(ws)
            c5_correct = (c5_check["status"] == "CORRECT")
            if c5_correct: score += 1.25
            details.append({'task': 'C5', 'correct': c5_correct, 'error': " | ".join(c5_check["issues"]) if not c5_correct else ""})

            c7_check = excel_auto_checker_skill2.check_basic_c7(ws, wb)
            c7_correct = (c7_check["status"] == "CORRECT")
            if c7_correct: score += 1.25
            details.append({'task': 'C7', 'correct': c7_correct, 'error': " | ".join(c7_check["issues"]) if not c7_correct else ""})
    except Exception as e:
        details.append({'error': str(e)})
    return round(min(score, 2.5), 4), details

def grade_dropdown_advanced(wb):
    score = 0; details = []
    try:
        for name, (sheet_name, cell_range) in excel_auto_checker_skill2.ADVANCED_NAMED_RANGES.items():
            check = excel_auto_checker_skill2.check_advanced_named_range(wb, name, sheet_name, cell_range)
            correct = (check["status"] == "CORRECT")
            if correct:
                score += excel_auto_checker_skill2.ADVANCED_DROPDOWN_MARKS / 4
            details.append({
                'task': f'Named Range - {name}',
                'correct': correct,
                'error': " | ".join(check["issues"]) if not correct else ""
            })

        if "DROPDOWN ADVANCED" not in wb.sheetnames:
            details.append({'task': 'C10', 'correct': False, 'error': "Sheet 'DROPDOWN ADVANCED' not found."})
            details.append({'task': 'C11', 'correct': False, 'error': "Sheet 'DROPDOWN ADVANCED' not found."})
        else:
            ws = wb["DROPDOWN ADVANCED"]
            c10_check = excel_auto_checker_skill2.check_advanced_c10(ws)
            c10_correct = (c10_check["status"] == "CORRECT")
            if c10_correct: score += 0.625
            details.append({'task': 'C10', 'correct': c10_correct, 'error': " | ".join(c10_check["issues"]) if not c10_correct else ""})

            c11_check = excel_auto_checker_skill2.check_advanced_c11(ws)
            c11_correct = (c11_check["status"] == "CORRECT")
            if c11_correct: score += 0.625
            details.append({'task': 'C11', 'correct': c11_correct, 'error': " | ".join(c11_check["issues"]) if not c11_correct else ""})
    except Exception as e:
        details.append({'error': str(e)})
    return round(min(score, 2.5), 4), details

def grade_workbook_validation(wb):
    score = 0; details = []
    try:
        if "DATA VALIDATION SKILL 2" not in wb.sheetnames:
            details.append({'task': 'C5 (Whole Number)', 'correct': False, 'error': "Sheet 'DATA VALIDATION SKILL 2' not found."})
            details.append({'task': 'C7 (Date)', 'correct': False, 'error': "Sheet 'DATA VALIDATION SKILL 2' not found."})
            details.append({'task': 'C9 (Text Length)', 'correct': False, 'error': "Sheet 'DATA VALIDATION SKILL 2' not found."})
        else:
            ws = wb["DATA VALIDATION SKILL 2"]
            c5_check = excel_auto_checker_skill2.check_whole_number_c5(ws)
            c5_correct = (c5_check["status"] == "CORRECT")
            if c5_correct: score += 0.8333
            details.append({'task': 'C5 (Whole Number)', 'correct': c5_correct, 'error': " | ".join(c5_check["issues"]) if not c5_correct else ""})

            c7_check = excel_auto_checker_skill2.check_date_c7(ws)
            c7_correct = (c7_check["status"] == "CORRECT")
            if c7_correct: score += 0.8333
            details.append({'task': 'C7 (Date)', 'correct': c7_correct, 'error': " | ".join(c7_check["issues"]) if not c7_correct else ""})

            c9_check = excel_auto_checker_skill2.check_text_length_c9(ws)
            c9_correct = (c9_check["status"] == "CORRECT")
            if c9_correct: score += 0.8334
            details.append({'task': 'C9 (Text Length)', 'correct': c9_correct, 'error': " | ".join(c9_check["issues"]) if not c9_correct else ""})
    except Exception as e:
        details.append({'error': str(e)})
    return round(min(score, 2.5), 4), details

def grade_vlookup(wb):
    score = 0; details = []
    try:
        ws = wb['VLOOKUP']
        # Q1: C19 (Find Department of Employee E003 -> Finance)
        v1 = ws['C19'].value
        v1_str = str(v1 or "").strip()
        if v1 and ('finance' in v1_str.lower() or ('vlookup' in v1_str.upper() and 'E003' in v1_str.upper())):
            score += 0.5; details.append({'q': 'Q1', 'task': 'Find Department of Employee E003 (VLOOKUP)', 'correct': True})
        else:
            details.append({'q': 'Q1', 'task': 'Find Department of Employee E003 (VLOOKUP)', 'correct': False, 'error': f"Expected 'Finance' or VLOOKUP formula for E003, Got '{v1}' (Cell C19)"})

        # Q2: C20 (Find Salary of Sara Khan -> 42000)
        v2 = ws['C20'].value
        v2_str = str(v2 or "").strip()
        if v2 and (v2_str in ['42000', '42000.0'] or ('vlookup' in v2_str.upper() and 'SARA' in v2_str.upper())):
            score += 0.5; details.append({'q': 'Q2', 'task': 'Find Salary of Sara Khan (VLOOKUP)', 'correct': True})
        else:
            details.append({'q': 'Q2', 'task': 'Find Salary of Sara Khan (VLOOKUP)', 'correct': False, 'error': f"Expected '42000' or VLOOKUP formula for Sara Khan, Got '{v2}' (Cell C20)"})

        # Q3: C21 (Find City of Employee E007 -> Islamabad)
        v3 = ws['C21'].value
        v3_str = str(v3 or "").strip()
        if v3 and ('islamabad' in v3_str.lower() or ('vlookup' in v3_str.upper() and 'E007' in v3_str.upper())):
            score += 0.5; details.append({'q': 'Q3', 'task': 'Find City of Employee E007 (VLOOKUP)', 'correct': True})
        else:
            details.append({'q': 'Q3', 'task': 'Find City of Employee E007 (VLOOKUP)', 'correct': False, 'error': f"Expected 'Islamabad' or VLOOKUP formula for E007, Got '{v3}' (Cell C21)"})

        # Q4: C22 (Find Name of Employee with ID E010 -> Maryam Fatima)
        v4 = ws['C22'].value
        v4_str = str(v4 or "").strip()
        if v4 and ('maryam' in v4_str.lower() or ('vlookup' in v4_str.upper() and 'E010' in v4_str.upper())):
            score += 0.5; details.append({'q': 'Q4', 'task': 'Find Name of Employee E010 (VLOOKUP)', 'correct': True})
        else:
            details.append({'q': 'Q4', 'task': 'Find Name of Employee E010 (VLOOKUP)', 'correct': False, 'error': f"Expected 'Maryam' or VLOOKUP formula for E010, Got '{v4}' (Cell C22)"})
    except Exception as e:
        details.append({'error': str(e)})
    return score, details

def grade_sumif_countif(wb):
    score = 0; details = []
    try:
        ws = wb['SUMIF & COUNTIF']
        q_checks = [
            {'cell': 'C20', 'q': 'Q5', 'task': 'Total sales by Ali (SUMIF)', 'expected': '157000', 'ans': ['157000', '157000.0'], 'fn': 'SUMIF'},
            {'cell': 'C21', 'q': 'Q6', 'task': 'Count of Laptop sales (COUNTIF)', 'expected': '4', 'ans': ['4', '4.0'], 'fn': 'COUNTIF'},
            {'cell': 'C22', 'q': 'Q7', 'task': 'Total quantity sold by Sara (SUMIF)', 'expected': '15', 'ans': ['15', '15.0'], 'fn': 'SUMIF'},
            {'cell': 'C23', 'q': 'Q8', 'task': 'Count of Electronics sold (COUNTIF)', 'expected': '7', 'ans': ['7', '7.0'], 'fn': 'COUNTIF'},
            {'cell': 'C24', 'q': 'Q9', 'task': 'Total amount of Accessories (SUMIF)', 'expected': '82000', 'ans': ['82000', '82000.0'], 'fn': 'SUMIF'},
            {'cell': 'C25', 'q': 'Q10', 'task': 'Count of Fatima sales (COUNTIF)', 'expected': '2', 'ans': ['2', '2.0'], 'fn': 'COUNTIF'}
        ]
        for qc in q_checks:
            v = ws[qc['cell']].value
            v_str = str(v or "").strip()
            if v and (v_str in qc['ans'] or qc['fn'] in v_str.upper()):
                score += 0.33; details.append({'q': qc['q'], 'task': qc['task'], 'correct': True})
            else:
                details.append({'q': qc['q'], 'task': qc['task'], 'correct': False, 'error': f"Expected value '{qc['expected']}' or {qc['fn']} formula, Got '{v}' (Cell {qc['cell']})"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2), details

def grade_text_functions(wb):
    score = 0; details = []
    try:
        ws = wb['LEFT RIGHT MID']
        q_checks = [
            {'cell': 'C15', 'q': 'Q11', 'task': 'Extract first 3 letters (LEFT)', 'expected': 'Ahm', 'check': lambda v, s: v and ('ahm' in s.lower() or 'left' in s.upper())},
            {'cell': 'C16', 'q': 'Q12', 'task': 'Extract last 7 digits (RIGHT)', 'expected': '1234567', 'check': lambda v, s: v and ('1234567' in s or 'right' in s.upper())},
            {'cell': 'C17', 'q': 'Q13', 'task': 'Extract username (MID/FIND)', 'expected': 'ahmed.ali', 'check': lambda v, s: v and ('ahmed' in s.lower() or 'mid' in s.upper() or 'find' in s.upper())},
            {'cell': 'C18', 'q': 'Q14', 'task': 'Extract year (RIGHT)', 'expected': '2024', 'check': lambda v, s: v and ('2024' in s or 'right' in s.upper())},
            {'cell': 'C19', 'q': 'Q15', 'task': 'Extract domain (MID)', 'expected': 'hotmail', 'check': lambda v, s: v and ('hotmail' in s.lower() or 'mid' in s.upper())},
            {'cell': 'C20', 'q': 'Q16', 'task': 'Extract middle name (MID)', 'expected': 'fatima', 'check': lambda v, s: v and ('fatima' in s.lower() or 'mid' in s.upper())}
        ]
        for qc in q_checks:
            v = ws[qc['cell']].value
            v_str = str(v or "").strip()
            if qc['check'](v, v_str):
                score += 0.33; details.append({'q': qc['q'], 'task': qc['task'], 'correct': True})
            else:
                details.append({'q': qc['q'], 'task': qc['task'], 'correct': False, 'error': f"Expected '{qc['expected']}' or text formula, Got '{v}' (Cell {qc['cell']})"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2), details

def grade_if_nested(wb):
    score = 0; details = []
    try:
        ws = wb['IF & NESTED IF']
        students = [
            {'row': 4, 'id': 'S001', 'name': 'Ahmed', 'g': 'A+', 's': 'Pass'},
            {'row': 5, 'id': 'S002', 'name': 'Sara', 'g': 'A', 's': 'Pass'},
            {'row': 6, 'id': 'S003', 'name': 'Omar', 'g': 'B', 's': 'Pass'},
            {'row': 7, 'id': 'S004', 'name': 'Fatima', 'g': 'C', 's': 'Pass'},
            {'row': 8, 'id': 'S005', 'name': 'Bilal', 'g': 'D', 's': 'Pass'},
            {'row': 9, 'id': 'S006', 'name': 'Ayesha', 'g': 'F', 's': 'Fail'},
            {'row': 10, 'id': 'S007', 'name': 'Hassan', 'g': 'F', 's': 'Fail'},
            {'row': 11, 'id': 'S008', 'name': 'Zainab', 'g': 'A', 's': 'Pass'},
            {'row': 12, 'id': 'S009', 'name': 'Ali', 'g': 'B', 's': 'Pass'},
            {'row': 13, 'id': 'S010', 'name': 'Maryam', 'g': 'A+', 's': 'Pass'}
        ]
        for std in students:
            g_cell = ws.cell(row=std['row'], column=4).value
            s_cell = ws.cell(row=std['row'], column=5).value
            g = str(g_cell or "").strip().upper()
            s = str(s_cell or "").strip().lower()
            if (g == std['g'] or 'if' in g.lower()) and (s == std['s'].lower() or 'if' in s.lower()):
                score += 0.2; details.append({'q': f"Student {std['id']}", 'task': f"Grade & Status for {std['name']}", 'correct': True})
            else:
                details.append({'q': f"Student {std['id']}", 'task': f"Grade & Status for {std['name']}", 'correct': False, 'error': f"Expected Grade '{std['g']}' and Status '{std['s']}', Got Grade '{g_cell}' and Status '{s_cell}'"})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2), details

def grade_complex(wb):
    score = 0; details = []
    try:
        ws = wb['COMPLEX CHALLENGE']
        q_checks = [
            {'cell': 'C20', 'q': 'Q21', 'task': 'Extract category code from A4', 'expected': 'LAP', 'check': lambda v, s: v and ('lap' in s.lower() or 'right' in s.upper() or 'mid' in s.upper())},
            {'cell': 'C21', 'q': 'Q22', 'task': 'Count products in Accessories', 'expected': '5', 'check': lambda v, s: v and (s in ['5', '5.0'] or 'countif' in s.upper())},
            {'cell': 'C22', 'q': 'Q23', 'task': 'Total stock of Electronics', 'expected': '25', 'check': lambda v, s: v and (s in ['25', '25.0'] or 'sumif' in s.upper())},
            {'cell': 'C23', 'q': 'Q24', 'task': 'Find price of PRD-KEY-003', 'expected': '4500', 'check': lambda v, s: v and (s in ['4500', '4500.0'] or 'vlookup' in s.upper())},
            {'cell': 'C24', 'q': 'Q25', 'task': 'Extract first word from B5', 'expected': 'Wireless', 'check': lambda v, s: v and ('wireless' in s.lower() or 'left' in s.upper() or 'find' in s.upper())}
        ]
        for qc in q_checks:
            v = ws[qc['cell']].value
            v_str = str(v or "").strip()
            if qc['check'](v, v_str):
                score += 0.2; details.append({'q': qc['q'], 'task': qc['task'], 'correct': True})
            else:
                details.append({'q': qc['q'], 'task': qc['task'], 'correct': False, 'error': f"Expected '{qc['expected']}' or formula, Got '{v}' (Cell {qc['cell']})"})
        
        for i in range(5):
            r = 25 + i
            q_num = f"Q{26+i}"
            v = ws.cell(row=r, column=3).value
            if v is not None and str(v).strip() != '':
                score += 0.2; details.append({'q': q_num, 'task': f'Complex Challenge Task {26+i}', 'correct': True})
            else:
                details.append({'q': q_num, 'task': f'Complex Challenge Task {26+i}', 'correct': False, 'error': f"Cell C{r} is empty. Provide the required formula or value."})
    except Exception as e:
        details.append({'error': str(e)})
    return min(score, 2), details

def split_formula_args(formula):
    formula = str(formula).strip()
    if formula.startswith("="):
        formula = formula[1:]
    start = formula.find("(")
    end = formula.rfind(")")
    if start == -1 or end == -1:
        return []
    inside = formula[start + 1:end]
    args = []
    current = []
    depth = 0
    inside_string = False
    for char in inside:
        if char == '"':
            inside_string = not inside_string
            current.append(char)
            continue
        if not inside_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
            elif char == "," and depth == 0:
                args.append("".join(current).strip())
                current = []
                continue
        current.append(char)
    if current:
        args.append("".join(current).strip())
    return args

def get_function_name(formula):
    if not isinstance(formula, str) or not formula.startswith("="):
        return None
    match = re.match(r"=\s*([A-Z][A-Z0-9_]*)\s*\(", formula.upper())
    return match.group(1) if match else None

def clean_arg(arg):
    return str(arg).strip().replace("$", "").upper()

def equivalent_exact_match(arg):
    arg = clean_arg(arg)
    return arg in ("FALSE", "0")

def parse_range(arg):
    arg = clean_arg(arg)
    match = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", arg)
    if not match:
        return None
    return {
        "start_col": match.group(1),
        "start_row": int(match.group(2)),
        "end_col": match.group(3),
        "end_row": int(match.group(4)),
    }

def range_equals(arg, expected):
    parsed = parse_range(arg)
    if not parsed:
        return False
    return (
        parsed["start_col"] == expected["start_col"]
        and parsed["start_row"] == expected["start_row"]
        and parsed["end_col"] == expected["end_col"]
        and parsed["end_row"] == expected["end_row"]
    )

def check_q1(formula):
    result = {
        "marks": 0,
        "status": "WRONG",
        "issues": [],
        "checks": {}
    }

    # ------------------------------------
    # 1. Formula check
    # ------------------------------------
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result

    function = get_function_name(formula)

    result["checks"]["function"] = function == "VLOOKUP"

    if function != "VLOOKUP":
        result["issues"].append(
            "VLOOKUP function is required."
        )
        return result

    args = split_formula_args(formula)

    if len(args) < 3:
        result["issues"].append(
            "VLOOKUP does not contain enough arguments."
        )
        return result

    # ------------------------------------
    # 2. Lookup value
    # ------------------------------------
    lookup_value = clean_arg(args[0])

    lookup_correct = lookup_value in (
        '"E050"',
        "E050"
    )

    result["checks"]["lookup_value"] = lookup_correct

    if not lookup_correct:
        result["issues"].append(
            f"Wrong lookup value: {args[0]}. "
            f"Expected E050."
        )

    # ------------------------------------
    # 3. Lookup table
    #
    # Both are valid:
    # A3:E103  -> header included
    # A4:E103  -> data only
    #
    # $ references are also accepted
    # because parse_range() uses clean_arg()
    # ------------------------------------
    parsed_table = parse_range(args[1])

    valid_q1_ranges = [
        {
            "start_col": "A",
            "start_row": 3,
            "end_col": "E",
            "end_row": 103
        },
        {
            "start_col": "A",
            "start_row": 4,
            "end_col": "E",
            "end_row": 103
        }
    ]

    table_correct = (
        parsed_table is not None
        and any(
            parsed_table == valid_range
            for valid_range in valid_q1_ranges
        )
    )

    result["checks"]["lookup_table"] = table_correct

    if not table_correct:
        result["issues"].append(
            f"Wrong lookup table: {args[1]}. "
            f"Expected A3:E103 or A4:E103."
        )

    # ------------------------------------
    # 4. Column index
    # ------------------------------------
    try:
        column_index = int(
            clean_arg(args[2])
        )
    except:
        column_index = None

    column_correct = column_index == 5

    result["checks"]["column_index"] = column_correct

    if not column_correct:
        result["issues"].append(
            f"Wrong column index: {args[2]}. "
            f"Expected 5."
        )

    # ------------------------------------
    # 5. Exact match
    #
    # FALSE and 0 both mean exact match
    # ------------------------------------
    if len(args) >= 4:

        exact_correct = equivalent_exact_match(
            args[3]
        )

        result["checks"]["exact_match"] = exact_correct

        if not exact_correct:
            result["issues"].append(
                f"Approximate match used ({args[3]}). "
                f"Use FALSE or 0 for exact match."
            )

    else:

        result["checks"]["exact_match"] = False

        result["issues"].append(
            "VLOOKUP match argument missing. "
            "Exact match FALSE/0 is required."
        )

    # ------------------------------------
    # 6. Final scoring
    # ------------------------------------
    if all(result["checks"].values()):

        result["marks"] = 1.25
        result["status"] = "CORRECT"

    return result

def check_q2(formula):
    result = {
        "marks": 0,
        "status": "WRONG",
        "issues": [],
        "checks": {}
    }

    # ------------------------------------
    # 1. Formula check
    # ------------------------------------
    if not is_formula(formula):
        result["issues"].append("No formula found.")
        return result

    function = get_function_name(formula)

    result["checks"]["function"] = function == "SUMIF"

    if function != "SUMIF":
        result["issues"].append(
            "SUMIF function is required."
        )
        return result

    args = split_formula_args(formula)

    if len(args) < 3:
        result["issues"].append(
            "SUMIF requires 3 arguments."
        )
        return result

    # ------------------------------------
    # 2. Criteria range = Department C
    #
    # Both valid:
    # C3:C103 -> header included
    # C4:C103 -> data only
    # ------------------------------------
    parsed_criteria_range = parse_range(
        args[0]
    )

    valid_criteria_ranges = [
        {
            "start_col": "C",
            "start_row": 3,
            "end_col": "C",
            "end_row": 103
        },
        {
            "start_col": "C",
            "start_row": 4,
            "end_col": "C",
            "end_row": 103
        }
    ]

    criteria_range_correct = (
        parsed_criteria_range is not None
        and any(
            parsed_criteria_range == valid_range
            for valid_range in valid_criteria_ranges
        )
    )

    result["checks"]["criteria_range"] = (
        criteria_range_correct
    )

    if not criteria_range_correct:
        result["issues"].append(
            f"Wrong criteria range: {args[0]}. "
            f"Expected C3:C103 or C4:C103."
        )

    # ------------------------------------
    # 3. Criteria = IT
    # ------------------------------------
    criteria = clean_arg(args[1])

    criteria_correct = criteria in (
        '"IT"',
        "IT"
    )

    result["checks"]["criteria"] = criteria_correct

    if not criteria_correct:
        result["issues"].append(
            f"Wrong criteria: {args[1]}. "
            f"Expected IT."
        )

    # ------------------------------------
    # 4. Sum range = Salary E
    #
    # Both valid:
    # E3:E103 -> header included
    # E4:E103 -> data only
    # ------------------------------------
    parsed_sum_range = parse_range(
        args[2]
    )

    valid_sum_ranges = [
        {
            "start_col": "E",
            "start_row": 3,
            "end_col": "E",
            "end_row": 103
        },
        {
            "start_col": "E",
            "start_row": 4,
            "end_col": "E",
            "end_row": 103
        }
    ]

    sum_range_correct = (
        parsed_sum_range is not None
        and any(
            parsed_sum_range == valid_range
            for valid_range in valid_sum_ranges
        )
    )

    result["checks"]["sum_range"] = (
        sum_range_correct
    )

    if not sum_range_correct:
        result["issues"].append(
            f"Wrong sum range: {args[2]}. "
            f"Expected E3:E103 or E4:E103."
        )

    # ------------------------------------
    # 5. Final scoring
    # ------------------------------------
    if all(result["checks"].values()):

        result["marks"] = 1.25
        result["status"] = "CORRECT"

    return result

def check_q3(formula):
    """
    Q3:
    Count of students from Karachi using COUNTIF.

    Valid examples:
        =COUNTIF(D4:D103,"Karachi")
        =COUNTIF(D3:D103,"Karachi")

    D3 = header "City", isliye D3:D103 bhi logically valid hai.

    Total = 1.25 marks
    """

    MAX_MARKS = 1.25

    result = {
        "status": "WRONG",
        "marks": 0,
        "formula": formula,
        "checks": {},
        "issues": []
    }

    # -----------------------------------------
    # 1. Formula hona chahiye
    # -----------------------------------------
    if not isinstance(formula, str) or not formula.startswith("="):
        result["issues"].append(
            "COUNTIF formula nahi mila."
        )
        return result

    # Normalize
    f = formula.upper().replace(" ", "")

    # -----------------------------------------
    # 2. COUNTIF function
    # -----------------------------------------
    countif_function = f.startswith("=COUNTIF(")

    result["checks"]["COUNTIF_function"] = countif_function

    if not countif_function:
        result["issues"].append(
            "COUNTIF function required hai."
        )
        return result

    # -----------------------------------------
    # 3. Arguments extract
    # -----------------------------------------
    inside = f[
        f.find("(") + 1:
        f.rfind(")")
    ]

    args = [
        x.strip()
        for x in inside.split(",")
    ]

    if len(args) != 2:
        result["issues"].append(
            "COUNTIF ke exactly 2 arguments hone chahiye."
        )
        return result

    range_part = args[0]
    criteria = args[1]

    # -----------------------------------------
    # 4. Range validation
    #
    # D3 = City header
    # D4:D103 = student data
    #
    # DONO ACCEPTED
    # -----------------------------------------
    valid_ranges = [
        "D3:D103",
        "D4:D103"
    ]

    correct_range = range_part in valid_ranges

    result["checks"]["city_range"] = correct_range

    if not correct_range:
        result["issues"].append(
            f"Wrong range: {range_part}. "
            "Expected D3:D103 or D4:D103."
        )

    # -----------------------------------------
    # 5. Karachi criteria
    #
    # Accept:
    # "KARACHI"
    # KARACHI
    # -----------------------------------------
    correct_criteria = criteria in [
        '"KARACHI"',
        "KARACHI"
    ]

    result["checks"]["karachi_criteria"] = correct_criteria

    if not correct_criteria:
        result["issues"].append(
            'Criteria "Karachi" hona chahiye.'
        )

    # -----------------------------------------
    # 6. Final grading
    # -----------------------------------------
    if (
        countif_function
        and correct_range
        and correct_criteria
    ):
        result["status"] = "CORRECT"
        result["marks"] = MAX_MARKS

    return result

def check_q4(formula):
    """
    Q4:
    Salary C108 > 45000  -> High
    Otherwise             -> Low

    Total = 1.25 marks
    """

    MAX_MARKS = 1.25

    result = {
        "status": "WRONG",
        "marks": 0,
        "formula": formula,
        "checks": {},
        "issues": []
    }

    if not isinstance(formula, str) or not formula.startswith("="):
        result["issues"].append(
            "IF formula nahi mila."
        )
        return result

    f = formula.upper().replace(" ", "")

    if_function = f.startswith("=IF(")

    result["checks"]["IF_function"] = if_function

    if not if_function:
        result["issues"].append(
            "IF function required hai."
        )
        return result

    inside = f[
        f.find("(") + 1:
        f.rfind(")")
    ]

    args = [
        x.strip()
        for x in inside.split(",")
    ]

    if len(args) != 3:
        result["issues"].append(
            "IF ke exactly 3 arguments hone chahiye."
        )
        return result

    condition = args[0]
    true_value = args[1]
    false_value = args[2]

    correct_condition = (
        condition == "C108>45000"
    )

    result["checks"]["salary_condition"] = correct_condition

    if not correct_condition:
        result["issues"].append(
            f"Wrong condition: {condition}. "
            "Expected C108>45000."
        )

    correct_true = (
        true_value in ['"HIGH"', "HIGH"]
    )

    result["checks"]["true_result"] = correct_true

    if not correct_true:
        result["issues"].append(
            'TRUE result "High" hona chahiye.'
        )

    correct_false = (
        false_value in ['"LOW"', "LOW"]
    )

    result["checks"]["false_result"] = correct_false

    if not correct_false:
        result["issues"].append(
            'FALSE result "Low" hona chahiye.'
        )

    if (
        if_function
        and correct_condition
        and correct_true
        and correct_false
    ):
        result["status"] = "CORRECT"
        result["marks"] = MAX_MARKS

    return result

def grade_vlookup_sumif_countif_if(wb):
    score = 0
    details = []
    try:
        ws = wb['EXCEL SKILL 5']
        
        # Q1
        f1 = ws['C108'].value
        res1 = check_q1(f1)
        score += res1['marks']
        issues_str1 = " | ".join(res1['issues']) if res1['issues'] else "Formula + result correct (1.25/1.25)"
        details.append({
            'q': 'Q1',
            'task': 'Find Salary of Employee E050 (VLOOKUP)',
            'correct': res1['status'] == 'CORRECT',
            'error': issues_str1 if res1['status'] != 'CORRECT' else 'Formula + result correct (1.25/1.25)'
        })
        
        # Q2
        f2 = ws['C109'].value
        res2 = check_q2(f2)
        score += res2['marks']
        issues_str2 = " | ".join(res2['issues']) if res2['issues'] else "Formula + result correct (1.25/1.25)"
        details.append({
            'q': 'Q2',
            'task': 'Total Salary of IT Department (SUMIF)',
            'correct': res2['status'] == 'CORRECT',
            'error': issues_str2 if res2['status'] != 'CORRECT' else 'Formula + result correct (1.25/1.25)'
        })
        
        # Q3
        f3 = ws['C110'].value
        res3 = check_q3(f3)
        score += res3['marks']
        issues_str3 = " | ".join(res3['issues']) if res3['issues'] else "Formula + result correct (1.25/1.25)"
        details.append({
            'q': 'Q3',
            'task': 'Count of students from Karachi (COUNTIF)',
            'correct': res3['status'] == 'CORRECT',
            'error': issues_str3 if res3['status'] != 'CORRECT' else 'Formula + result correct (1.25/1.25)'
        })
        
        # Q4
        f4 = ws['C111'].value
        res4 = check_q4(f4)
        score += res4['marks']
        issues_str4 = " | ".join(res4['issues']) if res4['issues'] else "Formula + result correct (1.25/1.25)"
        details.append({
            'q': 'Q4',
            'task': 'If Salary > 45000 then "High" else "Low" (IF)',
            'correct': res4['status'] == 'CORRECT',
            'error': issues_str4 if res4['status'] != 'CORRECT' else 'Formula + result correct (1.25/1.25)'
        })
        
    except Exception as e:
        details.append({'error': str(e)})
        
    return round(score, 2), details
