import os
import json
from datetime import datetime
from app import app, db, Student, ExcelSubmission, QuizSubmission, MidTerm, AssignmentSubmission, ExcelSkillsAssignment
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
            # Header Style (Bold, Dark Background, White Text)
            {
                'repeatCell': {
                    'range': {'sheetId': new_sheet_id, 'startRowIndex': 0, 'endRowIndex': 1},
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {'red': 0.12, 'green': 0.3, 'blue': 0.47},
                            'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'bold': True, 'fontSize': 11},
                            'horizontalAlignment': 'CENTER'
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            },
            # Freeze Header
            {
                'updateSheetProperties': {
                    'properties': {'sheetId': new_sheet_id, 'gridProperties': {'frozenRowCount': 1}},
                    'fields': 'gridProperties.frozenRowCount'
                }
            },
            # Alternating Colors (Banding)
            {
                'addBanding': {
                    'bandedRange': {
                        'range': {'sheetId': new_sheet_id, 'startRowIndex': 1, 'endRowIndex': 100},
                        'rowProperties': {
                            'headerColor': {'red': 0.12, 'green': 0.3, 'blue': 0.47},
                            'firstBandColor': {'red': 1, 'green': 1, 'blue': 1},
                            'secondBandColor': {'red': 0.95, 'green': 0.95, 'blue': 0.95}
                        }
                    }
                }
            },
            # Auto-resize columns
            {
                'autoResizeDimensions': {
                    'dimensions': {'sheetId': new_sheet_id, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': len(headers)}
                }
            }
        ]
        
        service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={'requests': requests}).execute()
        return new_sheet_id
    except Exception as e:
        print(f"Error creating professional sheet {sheet_name}: {e}")
        return None

def export_final_marks():
    """Calculate and export final marks to Google Sheets"""
    print("🚀 Starting Final Marks export...")
    service, sheet_id = get_sheets_service()
    if not service:
        return False, "Google Sheets not configured"
        
    with app.app_context():
        students = Student.query.all()
        excel_assignments = ExcelSkillsAssignment.query.all()
        
        # Prepare Final Marks Summary
        summary_headers = ['Rank', 'Student ID', 'Name', 'Excel Skills (Avg %)', 'Quizzes (Avg %)', 'Assignments (Avg %)', 'Total Avg %', 'Status']
        create_professional_sheet(service, sheet_id, 'Final Marks', summary_headers)
        
        # Prepare Individual Marks Details
        detail_headers = ['Student ID', 'Name', 'Category', 'Task/Title', 'Score', 'Total', 'Percentage', 'Date']
        create_professional_sheet(service, sheet_id, 'Individual Marks', detail_headers)
        
        summary_data = []
        detail_data = []
        
        for student in students:
            # 1. Excel Marks
            excel_subs = ExcelSubmission.query.filter_by(student_id=student.id).all()
            excel_pct = sum([sub.percentage for sub in excel_subs]) / len(excel_assignments) if excel_assignments else 0
            for sub in excel_subs:
                detail_data.append([student.student_id, student.name, 'Excel', sub.assignment.title, sub.score, 10, f"{sub.percentage}%", sub.submitted_at.strftime('%Y-%m-%d')])
                
            # 2. Quiz Marks
            quiz_subs = QuizSubmission.query.filter_by(student_id=student.id).all()
            quiz_pct = sum([sub.percentage for sub in quiz_subs]) / len(quiz_subs) if quiz_subs else 0
            for sub in quiz_subs:
                detail_data.append([student.student_id, student.name, 'Quiz', sub.quiz.title, sub.score, sub.total_questions, f"{sub.percentage}%", sub.submitted_at.strftime('%Y-%m-%d')])
                
            # 3. Assignment Marks
            assign_subs = AssignmentSubmission.query.filter_by(student_id=student.id).all()
            assign_pct = 0 # Simplified
            for sub in assign_subs:
                grade = sub.grade if sub.grade else 0
                detail_data.append([student.student_id, student.name, 'Assignment', sub.assignment.title, grade, 100, f"{grade}%", sub.submitted_at.strftime('%Y-%m-%d')])
            
            total_avg = (excel_pct + quiz_pct) / 2 # Weightage can be adjusted
            status = 'PASS' if total_avg >= 50 else 'FAIL'
            summary_data.append([0, student.student_id, student.name, f"{round(excel_pct, 1)}%", f"{round(quiz_pct, 1)}%", "-", f"{round(total_avg, 1)}%", status])
            
        # Sort summary by Total Avg %
        summary_data.sort(key=lambda x: float(x[6].strip('%')), reverse=True)
        for i, row in enumerate(summary_data, 1):
            row[0] = i # Set rank
            
        # Upload data
        if summary_data:
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id, range="'Final Marks'!A2",
                valueInputOption='USER_ENTERED', body={'values': summary_data}
            ).execute()
            
        if detail_data:
            service.spreadsheets().values().update(
                spreadsheetId=sheet_id, range="'Individual Marks'!A2",
                valueInputOption='USER_ENTERED', body={'values': detail_data}
            ).execute()
            
    print("✅ Final Marks and Individual Marks exported successfully!")
    return True, "Export successful"

if __name__ == "__main__":
    export_final_marks()
