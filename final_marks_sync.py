import os
import json
from datetime import datetime
from app import app, db, Student, ExcelSubmission, QuizSubmission, MidTerm, AssignmentSubmission, ExcelSkillsAssignment, Quiz
from clean_sheets_sync import get_sheets_service

def create_professional_sheet(service, sheet_id, sheet_name, headers):
    """Create a sheet with professional formatting"""
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
        sheets = [s['properties']['title'] for s in spreadsheet.get('sheets', [])]
        
        if sheet_name in sheets:
            # Delete existing to refresh
            sheet_obj = next(s for s in spreadsheet.get('sheets', []) if s['properties']['title'] == sheet_name)
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_id,
                body={'requests': [{'deleteSheet': {'sheetId': sheet_obj['properties']['sheetId']}}]}
            ).execute()
            
        # Add new sheet
        add_response = service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': [{'addSheet': {'properties': {'title': sheet_name}}}]}
        ).execute()
        
        new_sheet_id = add_response['replies'][0]['addSheet']['properties']['sheetId']
        
        # Add headers
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption='USER_ENTERED',
            body={'values': [headers]}
        ).execute()
        
        # Professional Formatting Requests
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
        print(f"Error creating sheet {sheet_name}: {e}")
        return None

def export_final_marks():
    """Calculate and export final marks with individual skill columns"""
    print("🚀 Starting Professional Final Marks export...")
    service, sheet_id = get_sheets_service()
    if not service:
        return False, "Google Sheets not configured"
        
    with app.app_context():
        students = Student.query.all()
        excel_assignments = ExcelSkillsAssignment.query.all()
        quizzes = Quiz.query.all()
        
        # Dynamic Headers
        summary_headers = ['Rank', 'Student ID', 'Name']
        for ea in excel_assignments:
            summary_headers.append(ea.title)
        
        summary_headers += ['Quizzes Avg %', 'Total Avg %', 'Status']
        
        create_professional_sheet(service, sheet_id, 'Final Marks', summary_headers)
        
        summary_data = []
        for student in students:
            row = [0, student.student_id, student.name]
            
            # 1. Excel Individual Skills
            total_excel_pct = 0
            for ea in excel_assignments:
                sub = ExcelSubmission.query.filter_by(student_id=student.id, assignment_id=ea.id).first()
                pct = sub.percentage if sub else 0
                row.append(f"{pct}%")
                total_excel_pct += pct
            
            excel_avg = total_excel_pct / len(excel_assignments) if excel_assignments else 0
            
            # 2. Quiz Average
            quiz_subs = QuizSubmission.query.filter_by(student_id=student.id).all()
            quiz_avg = sum([sub.percentage for sub in quiz_subs]) / len(quiz_subs) if quiz_subs else 0
            row.append(f"{round(quiz_avg, 1)}%")
            
            # 3. Total Average (Excel 50%, Quiz 50%)
            total_avg = (excel_avg + quiz_avg) / 2
            row.append(f"{round(total_avg, 1)}%")
            
            # 4. Status (Passing threshold 40% to be more lenient during course)
            status = 'PASS' if total_avg >= 40 else 'FAIL'
            row.append(status)
            
            summary_data.append(row)
            
        # Sort by Total Avg
        summary_data.sort(key=lambda x: float(x[-2].replace('%', '')), reverse=True)
        for i, row in enumerate(summary_data, 1):
            row[0] = i # Set rank
            
        if summary_data:
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id, range="'Final Marks'!A2",
                valueInputOption='USER_ENTERED', body={'values': summary_data}
            ).execute()
            
    return True, "Export successful"

def get_final_marks_from_sheet():
    """Fetch the Final Marks data with dynamic columns"""
    service, sheet_id = get_sheets_service()
    if not service: return None
    try:
        result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range="'Final Marks'!A:Z").execute()
        rows = result.get('values', [])
        if not rows or len(rows) < 2: return []
        
        headers = rows[0]
        data = []
        for row in rows[1:]:
            while len(row) < len(headers): row.append("-")
            data.append(dict(zip(headers, row)))
        return {'headers': headers, 'data': data}
    except Exception as e:
        print(f"Error: {e}")
        return None
