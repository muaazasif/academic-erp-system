import os
import json
import re
from datetime import datetime
from clean_sheets_sync import get_sheets_service

def get_assignment_max_marks(title):
    """Determine correct max marks: Skills 1-4 = 10, Skills 5-6 = 5"""
    if not title:
        return 10
    t = title.lower()
    if any(k in t for k in ['skill 5', 'skill 6', 'iferror', 'date', 'text']):
        return 5
    if any(k in t for k in ['skill 1', 'skill 2', 'skill 3', 'skill 4']):
        return 10
    return 10

def parse_existing_value(val_str, max_m):
    """Parse exact obtained marks directly from existing sheet string without halving 5-mark assignments"""
    if not val_str or val_str == '-' or str(val_str).strip() == '':
        return 0.0, 0.0
    try:
        val_str_lower = str(val_str).lower().strip()
        if '/' in val_str_lower:
            parts = val_str_lower.split('/')
            num = float(parts[0].strip())
            score = min(num, max_m)
            pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
            return score, pct
            
        if '%' in val_str_lower:
            m = re.search(r'([\d\.]+)\s*%', val_str_lower)
            if m:
                pct_val = float(m.group(1))
                if max_m == 5:
                    score = min(pct_val / 10.0, 5.0)
                else:
                    score = round((pct_val / 100) * max_m, 3)
                    score = min(score, max_m)
                pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
                return score, pct
        try:
            num = float(val_str_lower)
            score = min(num, max_m)
            pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
            return score, pct
        except: pass
    except Exception as e:
        print(f"Error parsing existing value '{val_str}': {e}")
    return 0.0, 0.0

def create_professional_sheet(service, sheet_id, sheet_name, headers):
    """Create or clear a sheet with professional formatting"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = {s['properties']['title']: s['properties']['sheetId'] for s in spreadsheet.get('sheets', [])}
        
        if sheet_name in sheets:
            if len(sheets) == 1:
                service.spreadsheets().values().clear(
                    spreadsheetId=sheet_id,
                    range=f"'{sheet_name}'!A:Z"
                ).execute()
                new_sheet_id = sheets[sheet_name]
            else:
                service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id,
                    body={'requests': [{'deleteSheet': {'sheetId': sheets[sheet_name]}}]}
                ).execute()
                add_response = service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id,
                    body={'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
                ).execute()
                new_sheet_id = add_response['replies'][0]['addSheet']['properties']['sheetId']
        else:
            add_response = service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
            ).execute()
            new_sheet_id = add_response['replies'][0]['addSheet']['properties']['sheetId']
        
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': [headers]}
        ).execute()
        
        requests = [
            {
                'repeatCell': {
                    'range': {'sheetId': new_sheet_id, 'startRowIndex': 0, 'endRowIndex': 1},
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {'red': 0.12, 'green': 0.3, 'blue': 0.47},
                            'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'bold': True},
                            'horizontalAlignment': 'CENTER'
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {'sheetId': new_sheet_id, 'gridProperties': {'frozenRowCount': 1}},
                    'fields': 'gridProperties.frozenRowCount'
                }
            },
            {
                'autoResizeDimensions': {
                    'dimensions': {'sheetId': new_sheet_id, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': len(headers)}
                }
            }
        ]
        service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={'requests': requests}).execute()
        return new_sheet_id
    except Exception as e:
        print(f"Error preparing sheet {sheet_name}: {e}")
        return None

def export_final_marks():
    """Calculate and export final marks with individual skill columns to 'Final Results' sheet, respecting 'Final Marks' sheet columns"""
    print("🚀 Starting Professional Final Marks export to 'Final Results' sheet...")
    from app import app, db, Student, ExcelSubmission, QuizSubmission, SQLSubmission, ExcelSkillsAssignment, Quiz, SQLSkillsAssignment
    service, sheet_id = get_sheets_service()
    if not service:
        return False, "Google Sheets not configured"
        
    with app.app_context():
        students = Student.query.all()
        excel_assignments = ExcelSkillsAssignment.query.all()
        sql_assignments = SQLSkillsAssignment.query.all()
        quizzes = Quiz.query.all()
        
        # 0. FETCH DATA FROM "SQL Assignments" SHEET (Merge Logic)
        sql_sheet_marks = {} # {(student_id, assignment_title): percentage}
        try:
            sql_res = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'SQL Assignments'!A:E").execute()
            sql_rows = sql_res.get('values', [])
            if len(sql_rows) > 1:
                for row in sql_rows[1:]:
                    if len(row) >= 5:
                        sid = str(row[0]).strip()
                        title = str(row[2]).strip()
                        pct_str = str(row[4]).replace('%', '').strip()
                        try:
                            pct = float(pct_str)
                            sql_sheet_marks[(sid, title)] = pct
                        except: pass
        except Exception as e:
            print(f"⚠️ Note: Could not read 'SQL Assignments' sheet for merging: {e}")

        # 0b. FETCH EXISTING HEADERS AND DATA FROM "Final Marks" SHEET (Respecting Column Changes)
        existing_final_marks = {} # {(student_id, header_name): value_str}
        sheet_assignment_headers = []
        try:
            existing_res = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'Final Marks'!A:Z").execute()
            existing_rows = existing_res.get('values', [])
            if len(existing_rows) > 1:
                h_idx = 0
                for idx, row in enumerate(existing_rows[:5]):
                    row_str = " ".join([str(c) for c in row]).lower()
                    if 'student id' in row_str or 'name' in row_str:
                        h_idx = idx
                        break
                ex_headers = [str(h).strip() for h in existing_rows[h_idx]]
                
                # Identify assignment columns in Final Marks (excluding Student ID, Name, Rank, Averages, Status)
                skip_cols = ['student id', 'name', 'rank', 'excel avg %', 'sql avg %', 'quizzes avg %', 'total avg %', 'status']
                sheet_assignment_headers = [h for h in ex_headers if h.lower() not in skip_cols and h]
                
                for r in existing_rows[h_idx+1:]:
                    while len(r) < len(ex_headers): r.append("-")
                    row_dict = dict(zip(ex_headers, r))
                    sid_val = str(row_dict.get('Student ID', '')).strip()
                    if sid_val:
                        for h, val in row_dict.items():
                            existing_final_marks[(sid_val, h.strip())] = val
        except Exception as e:
            print(f"⚠️ Note: Could not read existing 'Final Marks' sheet: {e}")

        # Determine summary_headers: Use sheet_assignment_headers if available, else DB assignments
        if sheet_assignment_headers:
            assignment_headers = sheet_assignment_headers
        else:
            assignment_headers = []
            for ea in excel_assignments: assignment_headers.append(ea.title)
            for sa in sql_assignments: assignment_headers.append(sa.title)

        summary_headers = ['Rank', 'Student ID', 'Name'] + assignment_headers + ['Excel Avg %', 'SQL Avg %', 'Quizzes Avg %', 'Total Avg %', 'Status']
        
        # Prepare the new sheet (Create or Clear 'Final Results')
        res = create_professional_sheet(service, sheet_id, 'Final Results', summary_headers)
        if res is None:
            return False, "Failed to prepare Final Results sheet"
        
        summary_data = []
        for student in students:
            sid_str = student.student_id
            row = [0, sid_str, student.name]
            
            # Map assignment headers to scores
            excel_assignment_titles = [ea.title for ea in excel_assignments]
            sql_assignment_titles = [sa.title for sa in sql_assignments]
            
            total_excel_pct = 0
            excel_count = 0
            total_sql_pct = 0
            sql_count = 0
            
            for h in assignment_headers:
                max_m = get_assignment_max_marks(h)
                score = 0
                pct = 0
                
                # Check DB Submissions first if title matches
                sub = None
                if h in excel_assignment_titles:
                    ea_obj = next((ea for ea in excel_assignments if ea.title == h), None)
                    if ea_obj:
                        sub = ExcelSubmission.query.filter_by(student_id=sid_str, assignment_id=ea_obj.id).first()
                elif h in sql_assignment_titles:
                    sa_obj = next((sa for sa in sql_assignments if sa.title == h), None)
                    if sa_obj:
                        sub = SQLSubmission.query.filter_by(student_id=sid_str, assignment_id=sa_obj.id).first()
                
                if sub and sub.score is not None:
                    score = float(sub.score)
                    score = min(score, max_m)
                    pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
                elif h in sql_sheet_marks and sid_str in [s for s, t in sql_sheet_marks.keys() if t == h]:
                    pct = sql_sheet_marks.get((sid_str, h), 0)
                    score = round((pct / 100) * max_m, 3)
                elif (sid_str, h.strip()) in existing_final_marks:
                    val_str = existing_final_marks[(sid_str, h.strip())]
                    score, pct = parse_existing_value(val_str, max_m)
                
                row.append(f"{score}/{max_m} ({pct}%)")
                
                if h in excel_assignment_titles or 'excel skill' in h.lower():
                    total_excel_pct += pct
                    excel_count += 1
                elif h in sql_assignment_titles or 'sql' in h.lower():
                    total_sql_pct += pct
                    sql_count += 1
            
            excel_avg = total_excel_pct / excel_count if excel_count > 0 else 0
            sql_avg = total_sql_pct / sql_count if sql_count > 0 else 0

            # Add Category Averages
            row.append(f"{round(excel_avg, 1)}%")
            row.append(f"{round(sql_avg, 1)}%")
            
            # Quiz Average
            quiz_subs = QuizSubmission.query.filter_by(student_id=sid_str).all()
            total_quiz_pct = 0
            for qsub in quiz_subs:
                total_q = len(qsub.quiz.questions) if qsub.quiz and qsub.quiz.questions else 0
                qscore = qsub.score if (qsub and qsub.score is not None) else 0
                try: qscore = float(qscore)
                except: qscore = 0
                if total_q > 0:
                    total_quiz_pct += (qscore / total_q) * 100
                
            quiz_avg = total_quiz_pct / len(quiz_subs) if quiz_subs else 0
            row.append(f"{round(quiz_avg, 1)}%")
            
            # Total Average
            categories = []
            if excel_count > 0: categories.append(excel_avg)
            if sql_count > 0: categories.append(sql_avg)
            if quizzes: categories.append(quiz_avg)
            
            total_avg = sum(categories) / len(categories) if categories else 0
            row.append(f"{round(total_avg, 1)}%")
            
            # Status
            status = 'PASS' if total_avg >= 70 else 'FAIL'
            row.append(status)
            
            summary_data.append(row)
            
        # Sort by Total Avg (Rank-wise sort)
        if summary_data:
            summary_data.sort(key=lambda x: float(str(x[-2]).replace('%', '')), reverse=True)
            for i, row in enumerate(summary_data, 1):
                row[0] = i # Set rank
            
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id, range="'Final Results'!A2",
                valueInputOption='USER_ENTERED', body={'values': summary_data}
            ).execute()
            
    return True, "Export successful"

def get_final_marks_from_sheet():
    """Fetch the Final Marks data from 'Final Results' sheet (with fallback to 'Final Marks')"""
    service, sheet_id = get_sheets_service()
    if not service: return None
    for sheet_name in ['Final Results', 'Final Marks']:
        try:
            result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=f"'{sheet_name}'!A:Z").execute()
            rows = result.get('values', [])
            if rows and len(rows) >= 2:
                header_idx = 0
                for idx, row in enumerate(rows[:5]):
                    row_str = " ".join([str(c) for c in row]).lower()
                    if 'student id' in row_str or 'name' in row_str or 'rank' in row_str:
                        header_idx = idx
                        break
                        
                headers = rows[header_idx]
                data = []
                for row in rows[header_idx+1:]:
                    while len(row) < len(headers): row.append("-")
                    data.append(dict(zip(headers, row)))
                return {'headers': headers, 'data': data}
        except Exception as e:
            print(f"Error fetching from {sheet_name}: {e}")
    return None

def get_student_fallback_record(target_student_id):
    """Generate final marks record directly from local DB for a single student if not in Google Sheets"""
    from app import app, db, Student, ExcelSubmission, QuizSubmission, SQLSubmission, ExcelSkillsAssignment, Quiz, SQLSkillsAssignment
    with app.app_context():
        student = Student.query.filter_by(student_id=str(target_student_id)).first()
        if not student:
            student = Student.query.filter_by(id=target_student_id).first()
        
        sid_str = str(target_student_id)
        name = student.name if student else f"Student {sid_str}"
        
        excel_assignments = ExcelSkillsAssignment.query.all()
        sql_assignments = SQLSkillsAssignment.query.all()
        quizzes = Quiz.query.all()
        
        headers = ['Rank', 'Student ID', 'Name']
        for ea in excel_assignments:
            headers.append(ea.title)
        for sa in sql_assignments:
            headers.append(sa.title)
        headers += ['Excel Avg %', 'SQL Avg %', 'Quizzes Avg %', 'Total Avg %', 'Status']
        
        row_dict = {
            'Rank': '-',
            'Student ID': sid_str,
            'Name': name
        }
        
        total_excel_pct = 0
        for ea in excel_assignments:
            sub = ExcelSubmission.query.filter_by(student_id=sid_str, assignment_id=ea.id).first()
            if not sub and student:
                sub = ExcelSubmission.query.filter_by(student_id=student.student_id, assignment_id=ea.id).first()
            max_m = get_assignment_max_marks(ea.title)
            
            score = 0
            pct = 0
            if sub and sub.score is not None:
                score = float(sub.score)
                score = min(score, max_m)
                pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
            
            row_dict[ea.title] = f"{score}/{max_m} ({pct}%)"
            total_excel_pct += pct
        excel_avg = total_excel_pct / len(excel_assignments) if excel_assignments else 0
        
        total_sql_pct = 0
        for sa in sql_assignments:
            sub = SQLSubmission.query.filter_by(student_id=sid_str, assignment_id=sa.id).first()
            if not sub and student:
                sub = SQLSubmission.query.filter_by(student_id=student.student_id, assignment_id=sa.id).first()
            score = sub.score if (sub and sub.score is not None) else 0
            max_m = sa.max_marks if (sa and sa.max_marks is not None) else 10
            pct = round((score / max_m) * 100, 1) if max_m > 0 else 0
            row_dict[sa.title] = f"{score}/{max_m} ({pct}%)"
            total_sql_pct += pct
        sql_avg = total_sql_pct / len(sql_assignments) if sql_assignments else 0
        
        row_dict['Excel Avg %'] = f"{round(excel_avg, 1)}%"
        row_dict['SQL Avg %'] = f"{round(sql_avg, 1)}%"
        
        quiz_subs = QuizSubmission.query.filter_by(student_id=sid_str).all()
        if not quiz_subs and student:
            quiz_subs = QuizSubmission.query.filter_by(student_id=student.student_id).all()
            
        total_quiz_pct = 0
        for sub in quiz_subs:
            total_q = len(sub.quiz.questions) if sub.quiz and sub.quiz.questions else 0
            score = sub.score if (sub and sub.score is not None) else 0
            try: score = float(score)
            except: score = 0
            if total_q > 0:
                total_quiz_pct += (score / total_q) * 100
        quiz_avg = total_quiz_pct / len(quiz_subs) if quiz_subs else 0
        row_dict['Quizzes Avg %'] = f"{round(quiz_avg, 1)}%"
        
        categories = []
        if excel_assignments: categories.append(excel_avg)
        if sql_assignments: categories.append(sql_avg)
        if quizzes: categories.append(quiz_avg)
        total_avg = sum(categories) / len(categories) if categories else 0
        row_dict['Total Avg %'] = f"{round(total_avg, 1)}%"
        row_dict['Status'] = 'PASS' if total_avg >= 70 else 'FAIL'
        
        return {
            'headers': headers,
            'data': [row_dict]
        }
