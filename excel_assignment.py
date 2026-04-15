"""
Excel Skills Assignment Generator & Auto-Grader
Creates exercises: VLOOKUP, SUMIF, COUNTIF, LEFT, RIGHT, MID (Easy → Hard)
Auto-grades student submissions out of 10 marks
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from io import BytesIO
import re
import json

# ============================================
# EXERCISE GENERATOR
# ============================================

def create_excel_exercise_workbook():
    """Create a complete Excel exercise workbook with all exercises"""
    wb = openpyxl.Workbook()
    
    # Remove default sheet
    wb.remove(wb.active)
    
    # Create sheets in order
    create_instructions_sheet(wb)
    create_vlookup_sheet(wb)
    create_sumif_countif_sheet(wb)
    create_text_functions_sheet(wb)
    create_complex_challenge_sheet(wb)
    
    return wb

def create_instructions_sheet(wb):
    """Create instructions sheet"""
    ws = wb.create_sheet("Instructions")
    
    # Styling
    title_font = Font(size=16, bold=True, color="FFFFFF")
    title_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(size=12, bold=True, color="2F5496")
    normal_font = Font(size=11)
    
    # Title
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = "📊 EXCEL SKILLS ASSIGNMENT - Instructions"
    title_cell.font = title_font
    title_cell.fill = title_fill
    title_cell.alignment = Alignment(horizontal='center')
    
    # Instructions
    instructions = [
        "",
        "📋 ASSIGNMENT OVERVIEW:",
        "Complete all exercises in this workbook. Total Marks: 10",
        "",
        "📝 EXERCISES (Easy → Hard):",
        "1. VLOOKUP Exercises (2 marks)",
        "2. SUMIF & COUNTIF Exercises (2 marks)",  
        "3. LEFT, RIGHT, MID Text Functions (3 marks)",
        "4. Complex Challenge (3 marks) - Combines all functions",
        "",
        "✅ HOW TO SUBMIT:",
        "- Download this file and complete all exercises",
        "- Save your work (DO NOT change sheet names or structure)",
        "- Upload the completed file",
        "",
        "🎯 GRADING CRITERIA:",
        "- Each correct answer = marks as specified",
        "- Formula must be correct (not just the value)",
        "- Auto-grading will check your work",
        "",
        "⏰ TIME LIMIT: 2 hours",
        "💡 TIP: Read instructions carefully in each sheet!"
    ]
    
    for i, text in enumerate(instructions, 2):
        ws.cell(row=i, column=1, value=text).font = normal_font
    
    # Column widths
    ws.column_dimensions['A'].width = 60
    ws.column_dimensions['B'].width = 20
    
    return ws

def create_vlookup_sheet(wb):
    """VLOOKUP exercises sheet"""
    ws = wb.create_sheet("VLOOKUP Exercises")
    
    # Styling
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(size=14, bold=True, color="2F5496")
    instruction_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    # Sample data table (A1:D11)
    ws.merge_cells('A1:D1')
    title = ws['A1']
    title.value = "Employee Database"
    title.font = title_font
    
    # Headers
    headers = ['Employee ID', 'Name', 'Department', 'Salary']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
    
    # Data
    employees = [
        ['E001', 'Ahmed Ali', 'IT', 45000],
        ['E002', 'Sara Khan', 'HR', 42000],
        ['E003', 'Omar Sheikh', 'Finance', 50000],
        ['E004', 'Fatima Noor', 'IT', 48000],
        ['E005', 'Bilal Ahmed', 'Sales', 38000],
        ['E006', 'Ayesha Malik', 'HR', 41000],
        ['E007', 'Hassan Raza', 'Finance', 52000],
        ['E008', 'Zainab Hussain', 'Sales', 36000],
    ]
    
    for i, emp in enumerate(employees):
        for j, val in enumerate(emp):
            cell = ws.cell(row=4+i, column=1+j, value=val)
            cell.alignment = Alignment(horizontal='center')
    
    # Questions section
    question_row = 14
    ws.merge_cells(f'A{question_row}:D{question_row}')
    q_title = ws.cell(row=question_row, column=1, value="📝 EXERCISES - Use VLOOKUP to find answers")
    q_title.font = Font(size=12, bold=True, color="C00000")
    
    # Instructions
    ws.merge_cells(f'A{question_row+1}:D{question_row+1}')
    instr = ws.cell(row=question_row+1, column=1, value="Write VLOOKUP formulas in column B to find the correct answers:")
    instr.font = Font(size=11)
    instr.fill = instruction_fill
    
    # Exercise headers
    ex_headers = ['Question #', 'Your Answer (Formula)', 'Expected Result', 'Marks']
    for col, header in enumerate(ex_headers, 1):
        cell = ws.cell(row=question_row+3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Questions
    questions = [
        ['Q1: Find Department of Employee E003', '', 'Finance', '0.5'],
        ['Q2: Find Salary of Sara Khan', '', '42000', '0.5'],
        ['Q3: Find Name of Employee E007', '', 'Hassan Raza', '0.5'],
        ['Q4: Find Department of Zainab Hussain', '', 'Sales', '0.5'],
    ]
    
    for i, q in enumerate(questions):
        for j, val in enumerate(q):
            cell = ws.cell(row=question_row+4+i, column=1+j, value=val)
            if j == 1:  # Formula column - make it yellow for user input
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    # Answer key (hidden in very last rows - for grading)
    answer_row = 25
    ws.cell(row=answer_row, column=1, value="ANSWER_KEY").font = Font(color="FFFFFF")
    ws.cell(row=answer_row, column=2, value=json.dumps({
        'B18': {'formula': 'VLOOKUP', 'lookup': 'E003', 'result': 'Finance'},
        'B19': {'formula': 'VLOOKUP', 'lookup': 'Sara Khan', 'result': 42000},
        'B20': {'formula': 'VLOOKUP', 'lookup': 'E007', 'result': 'Hassan Raza'},
        'B21': {'formula': 'VLOOKUP', 'lookup': 'Zainab Hussain', 'result': 'Sales'}
    }))
    
    # Column widths
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 10
    
    return ws

def create_sumif_countif_sheet(wb):
    """SUMIF and COUNTIF exercises"""
    ws = wb.create_sheet("SUMIF & COUNTIF")
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(size=14, bold=True, color="2F5496")
    
    # Sales data
    ws.merge_cells('A1:E1')
    title = ws['A1']
    title.value = "Monthly Sales Data"
    title.font = title_font
    
    headers = ['Date', 'Salesperson', 'Product', 'Quantity', 'Amount']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    # Data
    sales_data = [
        ['2024-01-05', 'Ali', 'Laptop', 2, 120000],
        ['2024-01-10', 'Sara', 'Mouse', 10, 15000],
        ['2024-01-15', 'Ali', 'Keyboard', 5, 25000],
        ['2024-02-01', 'Omar', 'Laptop', 3, 180000],
        ['2024-02-10', 'Sara', 'Monitor', 4, 80000],
        ['2024-02-15', 'Ali', 'Mouse', 8, 12000],
        ['2024-03-01', 'Omar', 'Keyboard', 6, 30000],
        ['2024-03-10', 'Fatima', 'Laptop', 2, 120000],
        ['2024-03-15', 'Sara', 'Laptop', 1, 60000],
        ['2024-03-20', 'Fatima', 'Monitor', 3, 60000],
    ]
    
    for i, row in enumerate(sales_data):
        for j, val in enumerate(row):
            cell = ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 17
    ws.merge_cells(f'A{q_row}:E{q_row}')
    ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use SUMIF and COUNTIF").font = Font(size=12, bold=True, color="C00000")
    
    ex_headers = ['Question', 'Your Formula', 'Expected Result', 'Marks']
    for col, header in enumerate(ex_headers, 1):
        cell = ws.cell(row=q_row+2, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    questions = [
        ['Q5: Total sales by Ali (SUMIF)', '', '', '0.5'],
        ['Q6: Count of Laptop sales (COUNTIF)', '', '', '0.5'],
        ['Q7: Total quantity sold by Sara (SUMIF)', '', '', '0.5'],
        ['Q8: Count of Mouse sales (COUNTIF)', '', '', '0.5'],
    ]
    
    for i, q in enumerate(questions):
        for j, val in enumerate(q):
            cell = ws.cell(row=q_row+3+i, column=1+j, value=val)
            if j == 1:
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 10
    
    return ws

def create_text_functions_sheet(wb):
    """LEFT, RIGHT, MID exercises"""
    ws = wb.create_sheet("LEFT, RIGHT, MID")
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    # Data
    ws.merge_cells('A1:B1')
    ws['A1'].value = "Student Registration Data"
    ws['A1'].font = Font(size=14, bold=True, color="2F5496")
    
    ws['A3'] = 'Full Name'
    ws['B3'] = 'Phone Number'
    ws['C3'] = 'Email'
    for col in range(1, 4):
        ws.cell(row=3, column=col).font = header_font
        ws.cell(row=3, column=col).fill = header_fill
    
    students = [
        ['Ahmed Ali Khan', '0300-1234567', 'ahmed.ali@gmail.com'],
        ['Sara Fatima', '0321-7654321', 'sara.f@hotmail.com'],
        ['Muhammad Omar', '0333-9876543', 'omar.pk@yahoo.com'],
        ['Fatima Noor', '0345-1112233', 'fatima.noor@outlook.com'],
        ['Bilal Hussain', '0301-4445566', 'bilal.h@gmail.com'],
    ]
    
    for i, student in enumerate(students):
        for j, val in enumerate(student):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 12
    ws.merge_cells(f'A{q_row}:D{q_row}')
    ws.cell(row=q_row, column=1, value="📝 EXERCISES - Use LEFT, RIGHT, MID functions").font = Font(size=12, bold=True, color="C00000")
    
    ex_headers = ['Question', 'Your Formula', 'Expected Result', 'Marks']
    for col, header in enumerate(ex_headers, 1):
        cell = ws.cell(row=q_row+2, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    questions = [
        ['Q9: Extract first 3 letters from A4 (LEFT)', '', '', '0.5'],
        ['Q10: Extract last 4 digits from B4 (RIGHT)', '', '', '0.5'],
        ['Q11: Extract username from C4 before @ (MID+FIND)', '', '', '0.5'],
        ['Q12: Extract phone code from B5 (LEFT+FIND)', '', '', '0.5'],
        ['Q13: Extract domain from C6 after @ (RIGHT+LEN+FIND)', '', '', '0.5'],
        ['Q14: Extract middle name from A5 (MID)', '', '', '0.5'],
    ]
    
    for i, q in enumerate(questions):
        for j, val in enumerate(q):
            cell = ws.cell(row=q_row+3+i, column=1+j, value=val)
            if j == 1:
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 10
    
    return ws

def create_complex_challenge_sheet(wb):
    """Complex challenge combining all functions"""
    ws = wb.create_sheet("Complex Challenge")
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    ws.merge_cells('A1:F1')
    ws['A1'].value = "🏆 CHALLENGE - Product Sales Analysis"
    ws['A1'].font = Font(size=14, bold=True, color="C00000")
    
    # Main data
    headers = ['Product Code', 'Product Name', 'Category', 'Price', 'Units Sold', 'Total Revenue']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    products = [
        ['PRD-001-LAP', 'Laptop Pro 15', 'Electronics', 85000, 12, None],
        ['PRD-002-MOU', 'Wireless Mouse', 'Accessories', 1500, 50, None],
        ['PRD-003-KEY', 'Mech Keyboard', 'Accessories', 4500, 25, None],
        ['PRD-004-MON', 'Monitor 27"', 'Electronics', 35000, 8, None],
        ['PRD-005-USB', 'USB Hub 7-in-1', 'Accessories', 2500, 40, None],
        ['PRD-006-HDD', 'External HDD 1TB', 'Storage', 8000, 15, None],
        ['PRD-007-SSD', 'SSD 500GB', 'Storage', 6500, 20, None],
        ['PRD-008-WEB', 'Webcam HD', 'Accessories', 3500, 30, None],
    ]
    
    for i, prod in enumerate(products):
        for j, val in enumerate(prod):
            ws.cell(row=4+i, column=1+j, value=val)
    
    # Questions
    q_row = 15
    ws.merge_cells(f'A{q_row}:F{q_row}')
    ws.cell(row=q_row, column=1, value="📝 CHALLENGE QUESTIONS (Use combinations of functions)").font = Font(size=12, bold=True, color="C00000")
    
    ex_headers = ['Question', 'Your Formula/Answer', 'Expected', 'Marks']
    for col, header in enumerate(ex_headers, 1):
        cell = ws.cell(row=q_row+2, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    
    challenges = [
        ['Q15: Calculate Total Revenue (Price × Units) for row 4', '', '', '0.5'],
        ['Q16: Extract category code from Product Code A4 (RIGHT+MID)', '', '', '0.5'],
        ['Q17: Count products in "Accessories" category (COUNTIF)', '', '', '0.5'],
        ['Q18: Total revenue of Electronics (SUMIF)', '', '', '0.5'],
        ['Q19: Find product name for PRD-003-KEY (VLOOKUP)', '', '', '0.5'],
        ['Q20: Extract first word from B5 (LEFT+FIND)', '', '', '0.5'],
    ]
    
    for i, q in enumerate(challenges):
        for j, val in enumerate(q):
            cell = ws.cell(row=q_row+3+i, column=1+j, value=val)
            if j == 1:
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 40
    ws.column_dimensions['C'].width = 15
    ws.column_dimensions['D'].width = 10
    
    return ws


# ============================================
# AUTO-GRADER
# ============================================

def grade_excel_submission(file_path):
    """Auto-grade a completed Excel assignment. Returns score out of 10"""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
    except Exception as e:
        return {'error': f'Cannot open file: {str(e)}', 'score': 0}
    
    total_score = 0
    max_score = 10
    details = {}
    
    # Grade VLOOKUP (2 marks)
    vlookup_score, vlookup_details = grade_vlookup(wb)
    total_score += vlookup_score
    details['VLOOKUP'] = {'score': vlookup_score, 'max': 2, 'details': vlookup_details}
    
    # Grade SUMIF/COUNTIF (2 marks)
    sumif_score, sumif_details = grade_sumif_countif(wb)
    total_score += sumif_score
    details['SUMIF/COUNTIF'] = {'score': sumif_score, 'max': 2, 'details': sumif_details}
    
    # Grade Text Functions (3 marks)
    text_score, text_details = grade_text_functions(wb)
    total_score += text_score
    details['Text Functions'] = {'score': text_score, 'max': 3, 'details': text_details}
    
    # Grade Complex Challenge (3 marks)
    challenge_score, challenge_details = grade_complex_challenge(wb)
    total_score += challenge_score
    details['Complex Challenge'] = {'score': challenge_score, 'max': 3, 'details': challenge_details}
    
    return {
        'score': round(min(total_score, 10), 2),
        'max': 10,
        'percentage': round((total_score / 10) * 100, 1),
        'details': details
    }

def grade_vlookup(wb):
    """Grade VLOOKUP exercises (2 marks total, 0.5 each)"""
    score = 0
    details = []
    
    try:
        ws = wb['VLOOKUP Exercises']
        
        # Check Q1: VLOOKUP E003 → Finance
        b18 = ws['B18']
        if b18.value and 'Finance' in str(b18.value):
            score += 0.5
            details.append({'q': 'Q1', 'correct': True})
        else:
            details.append({'q': 'Q1', 'correct': False})
        
        # Check Q2: VLOOKUP Sara Khan → 42000
        b19 = ws['B19']
        if b19.value and str(b19.value).strip() in ['42000', '42000.0']:
            score += 0.5
            details.append({'q': 'Q2', 'correct': True})
        else:
            details.append({'q': 'Q2', 'correct': False})
        
        # Check Q3: VLOOKUP E007 → Hassan Raza
        b20 = ws['B20']
        if b20.value and 'Hassan' in str(b20.value):
            score += 0.5
            details.append({'q': 'Q3', 'correct': True})
        else:
            details.append({'q': 'Q3', 'correct': False})
        
        # Check Q4: VLOOKUP Zainab → Sales
        b21 = ws['B21']
        if b21.value and 'Sales' in str(b21.value):
            score += 0.5
            details.append({'q': 'Q4', 'correct': True})
        else:
            details.append({'q': 'Q4', 'correct': False})
            
    except Exception as e:
        details.append({'error': str(e)})
    
    return score, details

def grade_sumif_countif(wb):
    """Grade SUMIF/COUNTIF exercises (2 marks total, 0.5 each)"""
    score = 0
    details = []
    
    try:
        ws = wb['SUMIF & COUNTIF']
        
        # Q5: Total sales by Ali = 120000+25000+12000 = 157000
        b20 = ws['B20']
        if b20.value and str(b20.value).strip() in ['157000', '157000.0']:
            score += 0.5
            details.append({'q': 'Q5', 'correct': True})
        else:
            details.append({'q': 'Q5', 'correct': False, 'expected': 157000})
        
        # Q6: Count of Laptop sales = 4
        b21 = ws['B21']
        if b21.value and str(b21.value).strip() in ['4', '4.0']:
            score += 0.5
            details.append({'q': 'Q6', 'correct': True})
        else:
            details.append({'q': 'Q6', 'correct': False, 'expected': 4})
        
        # Q7: Total quantity by Sara = 10+4+1 = 15
        b22 = ws['B22']
        if b22.value and str(b22.value).strip() in ['15', '15.0']:
            score += 0.5
            details.append({'q': 'Q7', 'correct': True})
        else:
            details.append({'q': 'Q7', 'correct': False, 'expected': 15})
        
        # Q8: Count of Mouse = 2
        b23 = ws['B23']
        if b23.value and str(b23.value).strip() in ['2', '2.0']:
            score += 0.5
            details.append({'q': 'Q8', 'correct': True})
        else:
            details.append({'q': 'Q8', 'correct': False, 'expected': 2})
            
    except Exception as e:
        details.append({'error': str(e)})
    
    return score, details

def grade_text_functions(wb):
    """Grade LEFT/RIGHT/MID exercises (3 marks total, 0.5 each)"""
    score = 0
    details = []
    
    try:
        ws = wb['LEFT, RIGHT, MID']
        
        # Q9: LEFT(A4,3) = "Ahm"
        b15 = ws['B15']
        if b15.value and str(b15.value).strip().lower() == 'ahm':
            score += 0.5
            details.append({'q': 'Q9', 'correct': True})
        else:
            details.append({'q': 'Q9', 'correct': False, 'expected': 'Ahm'})
        
        # Q10: RIGHT(B4,7) = "1234567"
        b16 = ws['B16']
        if b16.value and '1234567' in str(b16.value):
            score += 0.5
            details.append({'q': 'Q10', 'correct': True})
        else:
            details.append({'q': 'Q10', 'correct': False, 'expected': '1234567'})
        
        # Q11: Extract "ahmed.ali" from email
        b17 = ws['B17']
        if b17.value and 'ahmed' in str(b17.value).lower():
            score += 0.5
            details.append({'q': 'Q11', 'correct': True})
        else:
            details.append({'q': 'Q11', 'correct': False})
        
        # Q12: Extract "0333" from phone
        b18 = ws['B18']
        if b18.value and '0333' in str(b18.value):
            score += 0.5
            details.append({'q': 'Q12', 'correct': True})
        else:
            details.append({'q': 'Q12', 'correct': False, 'expected': '0333'})
        
        # Q13: Extract domain "yahoo.com"
        b19 = ws['B19']
        if b19.value and 'yahoo' in str(b19.value).lower():
            score += 0.5
            details.append({'q': 'Q13', 'correct': True})
        else:
            details.append({'q': 'Q13', 'correct': False})
        
        # Q14: Extract middle name "Omar"
        b20 = ws['B20']
        if b20.value and 'omar' in str(b20.value).lower():
            score += 0.5
            details.append({'q': 'Q14', 'correct': True})
        else:
            details.append({'q': 'Q14', 'correct': False, 'expected': 'Omar'})
            
    except Exception as e:
        details.append({'error': str(e)})
    
    return score, details

def grade_complex_challenge(wb):
    """Grade complex challenge (3 marks total, 0.5 each)"""
    score = 0
    details = []
    
    try:
        ws = wb['Complex Challenge']
        
        # Q15: 85000 * 12 = 1020000
        b18 = ws['B18']
        if b18.value and str(b18.value).strip() in ['1020000', '1020000.0']:
            score += 0.5
            details.append({'q': 'Q15', 'correct': True})
        else:
            details.append({'q': 'Q15', 'correct': False, 'expected': 1020000})
        
        # Q16: Extract "LAP" from PRD-001-LAP
        b19 = ws['B19']
        if b19.value and 'LAP' in str(b19.value).upper():
            score += 0.5
            details.append({'q': 'Q16', 'correct': True})
        else:
            details.append({'q': 'Q16', 'correct': False, 'expected': 'LAP'})
        
        # Q17: Count Accessories = 5
        b20 = ws['B20']
        if b20.value and str(b20.value).strip() in ['5', '5.0']:
            score += 0.5
            details.append({'q': 'Q17', 'correct': True})
        else:
            details.append({'q': 'Q17', 'correct': False, 'expected': 5})
        
        # Q18: Electronics total = 1020000 + 280000 = 1300000
        b21 = ws['B21']
        if b21.value and str(b21.value).strip() in ['1300000', '1300000.0']:
            score += 0.5
            details.append({'q': 'Q18', 'correct': True})
        else:
            details.append({'q': 'Q18', 'correct': False, 'expected': 1300000})
        
        # Q19: VLOOKUP PRD-003-KEY → Mech Keyboard
        b22 = ws['B22']
        if b22.value and 'keyboard' in str(b22.value).lower():
            score += 0.5
            details.append({'q': 'Q19', 'correct': True})
        else:
            details.append({'q': 'Q19', 'correct': False})
        
        # Q20: LEFT(B5,FIND(" ",B5)-1) → "Wireless"
        b23 = ws['B23']
        if b23.value and 'wireless' in str(b23.value).lower():
            score += 0.5
            details.append({'q': 'Q20', 'correct': True})
        else:
            details.append({'q': 'Q20', 'correct': False, 'expected': 'Wireless'})
            
    except Exception as e:
        details.append({'error': str(e)})
    
    return score, details
