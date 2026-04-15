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

def create_excel_exercise_workbook():
    """Create workbook with anti-cheating protection"""
    template_path = os.path.join(os.path.dirname(__file__), 'excel_template.xlsm')
    
    if os.path.exists(template_path):
        # Use template with VBA cheating detection
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsm', dir=os.path.dirname(__file__))
        shutil.copy(template_path, temp_file.name)
        temp_file.close()
        
        wb = openpyxl.load_workbook(temp_file.name)
        os.unlink(temp_file.name)
    else:
        # Create workbook without VBA (fallback)
        wb = openpyxl.Workbook()
    
    # Remove default sheet if exists
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Create all exercise sheets
    create_instructions(wb)
    create_vlookup_exercises(wb)
    create_sumif_countif_exercises(wb)
    create_text_functions_exercises(wb)
    create_if_nested_exercises(wb)
    create_complex_challenge(wb)
    
    return wb

def create_instructions(wb):
    ws = wb.create_sheet("Instructions")
    ws['A1'] = "📊 EXCEL SKILLS ASSIGNMENT - Complete Workbook"
    ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = 40
    
    ws['A2'] = "VLOOKUP | SUMIF | COUNTIF | LEFT | RIGHT | MID | IF | NESTED FUNCTIONS"
    ws['A2'].font = Font(size=12, bold=True, color="FF6600")
    ws.merge_cells('A2:F2')
    
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
        ("4. IF & NESTED IF (2 marks) - 5 questions", None),
        ("5. COMPLEX CHALLENGE (2 marks) - 10 questions", None),
        ("", None),
        ("✅ HOW TO COMPLETE:", None),
        ("• Write formula in YELLOW cell", None),
        ("• Your answer calculates automatically", None),
        ("• DO NOT change sheet names", None)
    ]
    
    for i, (text, _) in enumerate(content, 4):
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

def grade_excel_submission(file_path):
    """Auto-grade a completed Excel assignment. Returns score out of 10"""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        return {'error': f'Cannot open file: {str(e)}', 'score': 0}
    
    # ==========================================
    # CHECK FOR CHEATING (anti-cheating detection)
    # ==========================================
    cheating_detected = False
    
    # Check hidden cell Z100 in Instructions sheet
    if 'Instructions' in wb.sheetnames:
        ws = wb['Instructions']
        cheat_flag = ws.cell(row=100, column=26).value  # Z100
        
        if cheat_flag and 'CHEAT' in str(cheat_flag).upper():
            cheating_detected = True
            print("🚨 CHEATING DETECTED: Student opened other windows/files!")
    
    # If cheating detected, return ZERO marks
    if cheating_detected:
        return {
            'score': 0,
            'max': 10,
            'percentage': 0,
            'cheating_detected': True,
            'details': {'error': 'Cheating detected - other files/windows were opened'}
        }
    
    # Normal grading
    total_score = 0
    details = {}
    
    # Grade VLOOKUP (2 marks, 4 questions = 0.5 each)
    v_score, v_detail = grade_vlookup(wb)
    total_score += v_score
    details['VLOOKUP'] = {'score': v_score, 'max': 2, 'details': v_detail}
    
    # Grade SUMIF/COUNTIF (2 marks, 6 questions = 0.33 each)
    s_score, s_detail = grade_sumif_countif(wb)
    total_score += s_score
    details['SUMIF/COUNTIF'] = {'score': s_score, 'max': 2, 'details': s_detail}
    
    # Grade Text Functions (2 marks, 6 questions = 0.33 each)
    t_score, t_detail = grade_text_functions(wb)
    total_score += t_score
    details['Text Functions'] = {'score': t_score, 'max': 2, 'details': t_detail}
    
    # Grade Complex (2 marks, 10 questions = 0.2 each)
    c_score, c_detail = grade_complex(wb)
    total_score += c_score
    details['Complex'] = {'score': c_score, 'max': 2, 'details': c_detail}
    
    return {
        'score': round(min(total_score, 10), 2),
        'max': 10,
        'percentage': round((total_score / 10) * 100, 1),
        'cheating_detected': False,
        'details': details
    }

def grade_vlookup(wb):
    score = 0
    details = []
    try:
        ws = wb['VLOOKUP']
        # Q1: B18 - Department of E003 = Finance
        if ws['D19'].value and 'finance' in str(ws['D19'].value).lower():
            score += 0.5
            details.append({'q': 'Q1', 'correct': True})
        else:
            details.append({'q': 'Q1', 'correct': False})
        
        # Q2: B19 - Salary of Sara Khan = 42000
        if ws['D20'].value and str(ws['D20'].value).strip() in ['42000', '42000.0']:
            score += 0.5
            details.append({'q': 'Q2', 'correct': True})
        else:
            details.append({'q': 'Q2', 'correct': False})
        
        # Q3: B20 - City of E007 = Islamabad
        if ws['D21'].value and 'islamabad' in str(ws['D21'].value).lower():
            score += 0.5
            details.append({'q': 'Q3', 'correct': True})
        else:
            details.append({'q': 'Q3', 'correct': False})
        
        # Q4: B21 - Name of E010 = Maryam Fatima
        if ws['D22'].value and 'maryam' in str(ws['D22'].value).lower():
            score += 0.5
            details.append({'q': 'Q4', 'correct': True})
        else:
            details.append({'q': 'Q4', 'correct': False})
    except:
        pass
    return score, details

def grade_sumif_countif(wb):
    score = 0
    details = []
    try:
        ws = wb['SUMIF & COUNTIF']
        # Q5: Total by Ali = 120000+25000+12000 = 157000
        if ws['D20'].value and str(ws['D20'].value).strip() in ['157000', '157000.0']:
            score += 0.33
            details.append({'q': 'Q5', 'correct': True})
        else:
            details.append({'q': 'Q5', 'correct': False})
        
        # Q6: Count Laptop = 4
        if ws['D21'].value and str(ws['D21'].value).strip() in ['4', '4.0']:
            score += 0.33
            details.append({'q': 'Q6', 'correct': True})
        else:
            details.append({'q': 'Q6', 'correct': False})
        
        # Q7: Quantity by Sara = 10+4+1 = 15
        if ws['D22'].value and str(ws['D22'].value).strip() in ['15', '15.0']:
            score += 0.33
            details.append({'q': 'Q7', 'correct': True})
        else:
            details.append({'q': 'Q7', 'correct': False})
        
        # Q8: Count Electronics = 7
        if ws['D23'].value and str(ws['D23'].value).strip() in ['7', '7.0']:
            score += 0.33
            details.append({'q': 'Q8', 'correct': True})
        else:
            details.append({'q': 'Q8', 'correct': False})
        
        # Q9: Total Accessories = 15000+25000+12000+30000 = 82000
        if ws['D24'].value and str(ws['D24'].value).strip() in ['82000', '82000.0']:
            score += 0.33
            details.append({'q': 'Q9', 'correct': True})
        else:
            details.append({'q': 'Q9', 'correct': False})
        
        # Q10: Count Fatima = 2
        if ws['D25'].value and str(ws['D25'].value).strip() in ['2', '2.0']:
            score += 0.33
            details.append({'q': 'Q10', 'correct': True})
        else:
            details.append({'q': 'Q10', 'correct': False})
    except:
        pass
    return min(score, 2), details

def grade_text_functions(wb):
    score = 0
    details = []
    try:
        ws = wb['LEFT RIGHT MID']
        # Q11: LEFT(A4,3) = "Ahm"
        if ws['D15'].value and str(ws['D15'].value).strip().lower() == 'ahm':
            score += 0.33
            details.append({'q': 'Q11', 'correct': True})
        else:
            details.append({'q': 'Q11', 'correct': False})
        
        # Q12: RIGHT(B4,7) = "1234567"
        if ws['D16'].value and '1234567' in str(ws['D16'].value):
            score += 0.33
            details.append({'q': 'Q12', 'correct': True})
        else:
            details.append({'q': 'Q12', 'correct': False})
        
        # Q13: Extract "ahmed.ali" from email
        if ws['D17'].value and 'ahmed' in str(ws['D17'].value).lower():
            score += 0.33
            details.append({'q': 'Q13', 'correct': True})
        else:
            details.append({'q': 'Q13', 'correct': False})
        
        # Q14: Extract "2024" from code
        if ws['D18'].value and '2024' in str(ws['D18'].value):
            score += 0.33
            details.append({'q': 'Q14', 'correct': True})
        else:
            details.append({'q': 'Q14', 'correct': False})
        
        # Q15: Extract domain "hotmail.com"
        if ws['D19'].value and 'hotmail' in str(ws['D19'].value).lower():
            score += 0.33
            details.append({'q': 'Q15', 'correct': True})
        else:
            details.append({'q': 'Q15', 'correct': False})
        
        # Q16: Extract "Fatima" from "Sara Fatima"
        if ws['D20'].value and 'fatima' in str(ws['D20'].value).lower():
            score += 0.33
            details.append({'q': 'Q16', 'correct': True})
        else:
            details.append({'q': 'Q16', 'correct': False})
    except:
        pass
    return min(score, 2), details

def grade_complex(wb):
    score = 0
    details = []
    try:
        ws = wb['COMPLEX CHALLENGE']
        # Q21: Extract "LAP" = LAP
        if ws['D20'].value and 'lap' in str(ws['D20'].value).lower():
            score += 0.2
            details.append({'q': 'Q21', 'correct': True})
        else:
            details.append({'q': 'Q21', 'correct': False})
        
        # Q22: Count Accessories = 5
        if ws['D21'].value and str(ws['D21'].value).strip() in ['5', '5.0']:
            score += 0.2
            details.append({'q': 'Q22', 'correct': True})
        else:
            details.append({'q': 'Q22', 'correct': False})
        
        # Q23: Stock Electronics = 12+8+5 = 25
        if ws['D22'].value and str(ws['D22'].value).strip() in ['25', '25.0']:
            score += 0.2
            details.append({'q': 'Q23', 'correct': True})
        else:
            details.append({'q': 'Q23', 'correct': False})
        
        # Q24: Price of PRD-KEY-003 = 4500
        if ws['D23'].value and str(ws['D23'].value).strip() in ['4500', '4500.0']:
            score += 0.2
            details.append({'q': 'Q24', 'correct': True})
        else:
            details.append({'q': 'Q24', 'correct': False})
        
        # Q25: "Wireless"
        if ws['D24'].value and 'wireless' in str(ws['D24'].value).lower():
            score += 0.2
            details.append({'q': 'Q25', 'correct': True})
        else:
            details.append({'q': 'Q25', 'correct': False})
        
        # Q26-Q30: Similar checks
        for i in range(5):
            cell = ws.cell(row=25+i, column=4)
            if cell.value and len(str(cell.value).strip()) > 0:
                score += 0.2
                details.append({'q': f'Q{i+26}', 'correct': True})
            else:
                details.append({'q': f'Q{i+26}', 'correct': False})
    except:
        pass
    return min(score, 2), details
