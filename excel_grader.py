import openpyxl
from openpyxl.utils import get_column_letter

def grade_excel_assignment(submission_path, solution_path, assignment_title):
    """
    Grades an Excel assignment by comparing key cells with a solution file.
    """
    # Excel Skill 5: Formula-based grading
    if 'Excel Skill 5' in assignment_title:
        wb_sub = openpyxl.load_workbook(submission_path, data_only=False)
        try:
            ws = wb_sub['EXCEL SKILL 5']
            score = 0
            feedback = []

            # --- STRICT FORMULA CHECKS ---
            # Column C is where yellow cells are (C108-C111)
            
            # Q1: VLOOKUP (C108)
            f1 = str(ws['C108'].value).upper().replace(" ", "")
            if 'VLOOKUP' in f1 and '"E050"' in f1 and 'A4:E103' in f1 and ',5,0' in f1: 
                score += 1.25; feedback.append({'q': 'Q1', 'task': 'VLOOKUP', 'correct': True})
            else: feedback.append({'q': 'Q1', 'task': 'VLOOKUP', 'correct': False, 'error': f"Expected: =VLOOKUP(\"E050\", A4:E103, 5, 0). Aapka: {ws['C108'].value}"})

            # Q2: SUMIF (C109)
            f2 = str(ws['C109'].value).upper().replace(" ", "")
            if 'SUMIF' in f2 and 'C4:C103' in f2 and '"IT"' in f2 and 'E4:E103' in f2: 
                score += 1.25; feedback.append({'q': 'Q2', 'task': 'SUMIF', 'correct': True})
            else: feedback.append({'q': 'Q2', 'task': 'SUMIF', 'correct': False, 'error': f"Expected: =SUMIF(C4:C103, \"IT\", E4:E103). Aapka: {ws['C109'].value}"})

            # Q3: COUNTIF (C110)
            f3 = str(ws['C110'].value).upper().replace(" ", "")
            if 'COUNTIF' in f3 and 'C4:C103' in f3 and '"KARACHI"' in f3: 
                score += 1.25; feedback.append({'q': 'Q3', 'task': 'COUNTIF', 'correct': True})
            else: feedback.append({'q': 'Q3', 'task': 'COUNTIF', 'correct': False, 'error': f"Expected: =COUNTIF(C4:C103, \"Karachi\"). Aapka: {ws['C110'].value}"})

            # Q4: IF (C111)
            f4 = str(ws['C111'].value).upper().replace(" ", "")
            if 'IF' in f4 and 'E108>45000' in f4 and '"HIGH"' in f4 and '"LOW"' in f4: 
                score += 1.25; feedback.append({'q': 'Q4', 'task': 'IF', 'correct': True})
            else: feedback.append({'q': 'Q4', 'task': 'IF', 'correct': False, 'error': f"Expected: =IF(E108 > 45000, \"High\", \"Low\"). Aapka: {ws['C111'].value}"})

            return min(score, 5.0), feedback
        finally:
            wb_sub.close()
            
    # Default value-based grading for other assignments
    wb_sub = openpyxl.load_workbook(submission_path, data_only=True)
    wb_sol = openpyxl.load_workbook(solution_path, data_only=True)
    
    try:
        sheet_sub = wb_sub.active
        sheet_sol = wb_sol.active
        
        score = 0
        feedback = [] # Will be list of dicts
        
        grading_configs = {
            'default': [
                {'cell': 'B10', 'label': 'Task 1 Result', 'weight': 1},
                {'cell': 'B11', 'label': 'Task 2 Result', 'weight': 1},
                {'cell': 'B12', 'label': 'Task 3 Result', 'weight': 1},
                {'cell': 'B13', 'label': 'Task 4 Result', 'weight': 2},
            ]
        }
        
        checks = grading_configs.get(assignment_title, grading_configs['default'])
        
        for check in checks:
            sub_val = sheet_sub[check['cell']].value
            sol_val = sheet_sol[check['cell']].value
            
            if str(sub_val).strip() == str(sol_val).strip():
                score += check['weight']
                feedback.append({'task': check['label'], 'correct': True})
            else:
                feedback.append({'task': check['label'], 'correct': False, 'error': f"Expected '{sol_val}', Got '{sub_val}'"})
                
        return score, feedback
    finally:
        wb_sub.close()
        wb_sol.close()
