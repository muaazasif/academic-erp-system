from flask import Flask, render_template, request, redirect, url_for, flash, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime
import os
from dotenv import load_dotenv
import pickle
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-change-this'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///erp_system.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)


class Admin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
        
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(100), nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    def set_password(self, password):
        self.password_hash = generate_password_hash(password)
        
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    student_id = db.Column(db.String(50), db.ForeignKey('student.student_id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    check_in_time = db.Column(db.Time)
    check_out_time = db.Column(db.Time)
    status = db.Column(db.String(20), default='absent')  # present, absent, late
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    student = db.relationship('Student', backref=db.backref('attendances', lazy=True))


class Quiz(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    created_by = db.Column(db.String(100), db.ForeignKey('admin.username'), nullable=False)  # Admin who created
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    due_date = db.Column(db.DateTime)
    is_active = db.Column(db.Boolean, default=True)

    # Relationship to admin who created the quiz
    admin = db.relationship('Admin', backref=db.backref('quizzes', lazy=True))


class QuizQuestion(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    quiz_id = db.Column(db.Integer, db.ForeignKey('quiz.id'), nullable=False)
    question_text = db.Column(db.Text, nullable=False)
    question_number = db.Column(db.Integer, nullable=False)  # Order of the question

    # Relationship
    quiz = db.relationship('Quiz', backref=db.backref('questions', lazy=True, order_by="QuizQuestion.question_number"))


class QuizOption(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    question_id = db.Column(db.Integer, db.ForeignKey('quiz_question.id'), nullable=False)
    option_text = db.Column(db.Text, nullable=False)
    is_correct = db.Column(db.Boolean, default=False)  # Only one option should be correct

    # Relationship
    question = db.relationship('QuizQuestion', backref=db.backref('options', lazy=True))


class QuizSubmission(db.Model):
    """Stores student's answers to quiz questions"""
    id = db.Column(db.Integer, primary_key=True)
    quiz_id = db.Column(db.Integer, db.ForeignKey('quiz.id'), nullable=False)
    student_id = db.Column(db.String(50), db.ForeignKey('student.student_id'), nullable=False)
    submitted_at = db.Column(db.DateTime, default=datetime.utcnow)
    score = db.Column(db.Integer)  # Total score for the quiz

    # Relationship
    quiz = db.relationship('Quiz', backref=db.backref('submissions', lazy=True))
    student = db.relationship('Student', backref=db.backref('quiz_submissions', lazy=True))


class QuizAnswer(db.Model):
    """Stores individual answers to quiz questions"""
    id = db.Column(db.Integer, primary_key=True)
    submission_id = db.Column(db.Integer, db.ForeignKey('quiz_submission.id'), nullable=False)
    question_id = db.Column(db.Integer, db.ForeignKey('quiz_question.id'), nullable=False)
    selected_option_id = db.Column(db.Integer, db.ForeignKey('quiz_option.id'), nullable=False)  # Student's answer
    is_correct = db.Column(db.Boolean, default=False)  # Whether the answer was correct

    # Relationships
    submission = db.relationship('QuizSubmission', backref=db.backref('answers', lazy=True))
    question = db.relationship('QuizQuestion', backref=db.backref('answers', lazy=True))
    selected_option = db.relationship('QuizOption')


class Assignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    file_url = db.Column(db.String(500))  # Google Drive file URL
    created_by = db.Column(db.String(100), db.ForeignKey('admin.username'), nullable=False)  # Admin who created
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    due_date = db.Column(db.DateTime)
    is_active = db.Column(db.Boolean, default=True)

    # Relationship to admin who created the assignment
    admin = db.relationship('Admin', backref=db.backref('assignments', lazy=True))


class QuizAssignment(db.Model):
    """Link table to assign quizzes to students"""
    id = db.Column(db.Integer, primary_key=True)
    quiz_id = db.Column(db.Integer, db.ForeignKey('quiz.id'), nullable=False)
    student_id = db.Column(db.String(50), db.ForeignKey('student.student_id'), nullable=False)
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='assigned')  # assigned, submitted, graded
    grade = db.Column(db.Float)  # Numeric grade for the quiz

    # Relationships
    quiz = db.relationship('Quiz', backref=db.backref('quiz_assignments', lazy=True))
    student = db.relationship('Student', backref=db.backref('quiz_assignments', lazy=True))


class MidTerm(db.Model):
    """Mid-term exam with Excel workbook distribution"""
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    file_url = db.Column(db.String(500))  # URL to the original Excel workbook
    total_sheets = db.Column(db.Integer, nullable=False)  # Total sheets in the original workbook
    sheets_per_student = db.Column(db.Integer, nullable=False, default=5)  # Number of sheets per student
    created_by = db.Column(db.String(100), db.ForeignKey('admin.username'), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    due_date = db.Column(db.DateTime)
    is_active = db.Column(db.Boolean, default=True)

    # Relationship to admin who created the mid-term
    admin = db.relationship('Admin', backref=db.backref('mid_terms', lazy=True))


class MidTermAssignment(db.Model):
    """Link table to assign mid-term sheets to students"""
    id = db.Column(db.Integer, primary_key=True)
    mid_term_id = db.Column(db.Integer, db.ForeignKey('mid_term.id'), nullable=False)
    student_id = db.Column(db.String(50), db.ForeignKey('student.student_id'), nullable=False)
    assigned_sheets = db.Column(db.Text)  # Comma-separated list of sheet numbers assigned
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='assigned')  # assigned, submitted, graded
    grade = db.Column(db.Float)  # Numeric grade for the mid-term
    submission_file_path = db.Column(db.String(500))  # Path to submitted file

    # Relationships
    mid_term = db.relationship('MidTerm', backref=db.backref('mid_term_assignments', lazy=True))
    student = db.relationship('Student', backref=db.backref('mid_term_assignments', lazy=True))


class AssignmentSubmission(db.Model):
    """Link table to track assignment submissions"""
    id = db.Column(db.Integer, primary_key=True)
    assignment_id = db.Column(db.Integer, db.ForeignKey('assignment.id'), nullable=False)
    student_id = db.Column(db.String(50), db.ForeignKey('student.student_id'), nullable=False)
    submission_url = db.Column(db.String(500))  # Google Drive submission URL
    submitted_at = db.Column(db.DateTime)
    assigned_at = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='assigned')  # assigned, submitted, graded
    grade = db.Column(db.Float)  # Numeric grade for the assignment

    # Relationships
    assignment = db.relationship('Assignment', backref=db.backref('assignment_submissions', lazy=True))
    student = db.relationship('Student', backref=db.backref('assignment_submissions', lazy=True))


class CourseOutline(db.Model):
    """Course outline with presentations and resources"""
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(200), nullable=False)
    description = db.Column(db.Text)
    outline_file_url = db.Column(db.String(500))  # Google Drive URL for course outline
    presentation_file_url = db.Column(db.String(500))  # Google Drive URL for presentation
    week_number = db.Column(db.Integer)  # Week number for the course outline
    created_by = db.Column(db.String(100), db.ForeignKey('admin.username'), nullable=False)  # Admin who created
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    is_active = db.Column(db.Boolean, default=True)

    # Relationship to admin who created the course outline
    admin = db.relationship('Admin', backref=db.backref('course_outlines', lazy=True))


# Google Sheets Integration
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service object"""
    creds = None
    from google.oauth2.service_account import Credentials as ServiceAccountCredentials

    # First, try environment variable (works in Railway and other cloud platforms)
    creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    if creds_json:
        import json
        try:
            # Parse the credentials JSON from environment variable
            # Handle potential newline characters in the private key
            credentials_info = json.loads(creds_json)

            # Fix the private key by replacing literal \n with actual newlines
            if 'private_key' in credentials_info:
                # Properly decode the private key by handling escaped newlines
                private_key = credentials_info['private_key']
                # Replace escaped newlines with actual newlines
                private_key = private_key.replace('\\n', '\n')
                credentials_info['private_key'] = private_key

            # Create credentials object from the service account info
            creds = ServiceAccountCredentials.from_service_account_info(
                credentials_info, scopes=SCOPES
            )
            print("Successfully loaded credentials from environment variable")
        except (ValueError, TypeError) as env_error:
            print(f"Error parsing service account credentials from environment: {env_error}")

    # If environment variable didn't work, try local credentials.json file
    if not creds and os.path.exists('credentials.json'):
        try:
            creds = ServiceAccountCredentials.from_service_account_file(
                'credentials.json', scopes=SCOPES
            )
            print("Successfully loaded credentials from local credentials.json file")
        except Exception as local_error:
            print(f"Error using local credentials.json file: {local_error}")

    # Fallback to token.pickle if available
    if not creds and os.path.exists('token.pickle'):
        try:
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
            print("Successfully loaded credentials from token.pickle")
        except Exception as token_error:
            print(f"Error loading credentials from token.pickle: {token_error}")

    # For service account, we don't need to refresh or use OAuth flow
    # If we have valid credentials, use them; otherwise, return None gracefully
    if not creds:
        print("No valid credentials found for Google Sheets API. Google Sheets integration will be disabled.")
        return None

    try:
        service = build('sheets', 'v4', credentials=creds)
        return service
    except Exception as e:
        print(f"Error building Google Sheets service: {e}")
        return None


def add_attendance_to_sheet(student_id, name, date, check_in, check_out, status):
    """Add attendance record to Google Sheet"""
    try:
        service = authenticate_google_sheets()

        # If authentication failed, return False but log the reason
        if not service:
            print("Google Sheets service not available, skipping sync")
            return False

        sheet = service.spreadsheets()

        # Get the Google Sheet ID from environment variable
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID') or '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'

        if not SPREADSHEET_ID:
            print("Google Sheet ID not found in environment variables")
            return False

        # Prepare the data to insert
        values = [
            [str(date), student_id, name, str(check_in), str(check_out), status, str(datetime.now())]
        ]

        body = {
            'values': values
        }

        # Append the data to the sheet
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Attendance!A:H',  # Using 'Attendance' sheet tab
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f"Attendance data sent to Google Sheets successfully.")
        return True
    except Exception as e:
        print(f"Error adding to Google Sheets: {e}")
        return False


def add_assignment_submission_to_sheet(student_id, name, assignment_title, submission_url, submitted_at, grade=None):
    """Add assignment submission record to Google Sheet"""
    try:
        service = authenticate_google_sheets()
        sheet = service.spreadsheets()

        # Get the Google Sheet ID from environment variable
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID') or '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'

        if not SPREADSHEET_ID:
            print("Google Sheet ID not found in environment variables")
            return False

        # Prepare the data to insert
        values = [
            [student_id, name, assignment_title, submission_url, str(submitted_at), str(grade) if grade else "", str(datetime.now())]
        ]

        body = {
            'values': values
        }

        # Append the data to the sheet in a different range (Assignments sheet)
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Assignments!A:G',  # Using 'Assignments' sheet tab
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f"{result.get('updates').get('updatedCells')} cells updated in Assignments sheet.")
        return True
    except Exception as e:
        print(f"Error adding assignment submission to Google Sheets: {e}")
        return False


def add_quiz_submission_to_sheet(student_id, name, quiz_title, score, total_questions, submitted_at):
    """Add quiz submission record to Google Sheet"""
    try:
        service = authenticate_google_sheets()
        sheet = service.spreadsheets()

        # Get the Google Sheet ID from environment variable
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID') or '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'

        if not SPREADSHEET_ID:
            print("Google Sheet ID not found in environment variables")
            return False

        # Prepare the data to insert
        values = [
            [student_id, name, quiz_title, str(score), str(total_questions), f"{score}/{total_questions}", str(submitted_at), str(datetime.now())]
        ]

        body = {
            'values': values
        }

        # Append the data to the sheet in a different range (Quizzes sheet)
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='Quizzes!A:H',  # Using 'Quizzes' sheet tab
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f"{result.get('updates').get('updatedCells')} cells updated in Quizzes sheet.")
        return True
    except Exception as e:
        print(f"Error adding quiz submission to Google Sheets: {e}")
        return False


def add_detailed_quiz_answers_to_sheet(student_id, name, quiz_title, questions_data, submitted_at):
    """Add detailed quiz answers for each question to Google Sheet"""
    try:
        service = authenticate_google_sheets()
        sheet = service.spreadsheets()

        # Get the Google Sheet ID from environment variable
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID') or '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'

        if not SPREADSHEET_ID:
            print("Google Sheet ID not found in environment variables")
            return False

        # Prepare the data to be added - one row per question
        values = []
        for item in questions_data:
            question = item['question']
            selected_option = item['selected_option']
            correct_option = item['correct_option']
            is_correct = item['is_correct']

            values.append([
                student_id,
                name,
                quiz_title,
                question.question_number,
                question.question_text,
                selected_option.option_text if selected_option else 'No answer',
                correct_option.option_text if correct_option else '',
                'Correct' if is_correct else 'Incorrect',
                str(submitted_at),
                str(datetime.now())
            ])

        body = {
            'values': values
        }

        # Append the data to the sheet in a different range (Detailed Quiz Answers sheet)
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='DetailedQuizAnswers!A:J',  # Using 'DetailedQuizAnswers' sheet tab
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f"{result.get('updates').get('updatedCells')} cells updated in DetailedQuizAnswers sheet.")
        return True
    except Exception as e:
        print(f"Error adding detailed quiz answers to Google Sheets: {e}")
        return False


def add_midterm_grade_to_sheet(student_id, name, midterm_title, grade, graded_at):
    """Add midterm grade record to Google Sheet"""
    try:
        service = authenticate_google_sheets()
        sheet = service.spreadsheets()

        # Get the Google Sheet ID from environment variable
        SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID') or '1kRoHe5BFJG-Y2xPr29deuI79exsqgeGBl--gQOPAf20'

        if not SPREADSHEET_ID:
            print("Google Sheet ID not found in environment variables")
            return False

        # Prepare the data to insert
        values = [
            [student_id, name, midterm_title, str(grade), str(graded_at), str(datetime.now())]
        ]

        body = {
            'values': values
        }

        # Append the data to the sheet in a different range (Midterm Grades sheet)
        result = sheet.values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='MidtermGrades!A:F',  # Using 'MidtermGrades' sheet tab
            valueInputOption='RAW',
            body=body
        ).execute()

        print(f"{result.get('updates').get('updatedCells')} cells updated in MidtermGrades sheet.")
        return True
    except Exception as e:
        print(f"Error adding midterm grade to Google Sheets: {e}")
        return False


@app.route('/')
def index():
    if 'admin_id' in session:
        return redirect(url_for('admin_dashboard'))
    elif 'student_id' in session:
        return redirect(url_for('student_dashboard'))
    else:
        return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user_type = request.form['user_type']
        
        if user_type == 'admin':
            user = Admin.query.filter_by(username=username).first()
            if user and user.check_password(password):
                session['admin_id'] = user.id
                session['username'] = user.username
                return redirect(url_for('admin_dashboard'))
            else:
                flash('Invalid admin credentials')
        elif user_type == 'student':
            user = Student.query.filter_by(student_id=username).first()
            if user and user.check_password(password):
                session['student_id'] = user.student_id
                session['student_name'] = user.name
                return redirect(url_for('student_dashboard'))
            else:
                flash('Invalid student credentials')
    
    return render_template('login.html')


@app.route('/admin/dashboard')
def admin_dashboard():
    if 'admin_id' not in session:
        return redirect(url_for('login'))
    
    students = Student.query.all()
    total_students = len(students)
    
    # Get today's attendance
    today = datetime.today().date()
    present_count = Attendance.query.filter(
        Attendance.date == today,
        Attendance.status == 'present'
    ).count()
    
    return render_template('admin_dashboard.html', 
                          students=students, 
                          total_students=total_students,
                          present_count=present_count)


@app.route('/admin/add_student', methods=['GET', 'POST'])
def add_student():
    if 'admin_id' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        student_id = request.form['student_id']
        name = request.form['name']
        password = request.form['password']
        
        # Check if student ID already exists
        existing_student = Student.query.filter_by(student_id=student_id).first()
        if existing_student:
            flash('Student ID already exists')
            return render_template('add_student.html')
        
        # Create new student
        student = Student(student_id=student_id, name=name)
        student.set_password(password)
        db.session.add(student)
        db.session.commit()
        
        flash('Student added successfully')
        return redirect(url_for('admin_dashboard'))
    
    return render_template('add_student.html')


@app.route('/admin/edit_student/<int:student_id>', methods=['GET', 'POST'])
def edit_student(student_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))
    
    student = Student.query.get_or_404(student_id)
    
    if request.method == 'POST':
        student.name = request.form['name']
        new_password = request.form['password']
        if new_password:  # Only update password if provided
            student.set_password(new_password)
        
        db.session.commit()
        flash('Student updated successfully')
        return redirect(url_for('admin_dashboard'))
    
    return render_template('edit_student.html', student=student)


@app.route('/admin/delete_student/<int:student_id>')
def delete_student(student_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))
    
    student = Student.query.get_or_404(student_id)
    db.session.delete(student)
    db.session.commit()
    
    flash('Student deleted successfully')
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/attendance_report')
def attendance_report():
    if 'admin_id' not in session:
        return redirect(url_for('login'))
    
    attendances = Attendance.query.join(Student).order_by(Attendance.date.desc()).all()
    return render_template('attendance_report.html', attendances=attendances)


@app.route('/student/dashboard')
def student_dashboard():
    if 'student_id' not in session:
        return redirect(url_for('login'))
    
    student_id = session['student_id']
    student = Student.query.filter_by(student_id=student_id).first()
    
    # Get today's attendance record
    today = datetime.today().date()
    attendance = Attendance.query.filter_by(
        student_id=student_id,
        date=today
    ).first()
    
    return render_template('student_dashboard.html', 
                          student=student, 
                          attendance=attendance)


@app.route('/student/attendance_action', methods=['POST'])
def attendance_action():
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']
    action = request.form['action']  # 'check_in' or 'check_out'
    today = datetime.today().date()

    # Get or create attendance record for today
    attendance = Attendance.query.filter_by(
        student_id=student_id,
        date=today
    ).first()

    if not attendance:
        attendance = Attendance(student_id=student_id, date=today)
        db.session.add(attendance)

    if action == 'check_in':
        if attendance.check_in_time is None:
            attendance.check_in_time = datetime.now().time()
            attendance.status = 'present'
        else:
            flash('Already checked in today')
            return redirect(url_for('student_dashboard'))
    elif action == 'check_out':
        if attendance.check_in_time and attendance.check_out_time is None:
            attendance.check_out_time = datetime.now().time()
        else:
            flash('Check in first or already checked out')
            return redirect(url_for('student_dashboard'))

    db.session.commit()

    # Get student info for Google Sheets
    student = Student.query.filter_by(student_id=student_id).first()

    # Add to Google Sheets
    success = add_attendance_to_sheet(
        student_id=student.student_id,
        name=student.name,
        date=attendance.date,
        check_in=attendance.check_in_time,
        check_out=attendance.check_out_time,
        status=attendance.status
    )

    if not success:
        flash('Attendance recorded locally but failed to sync to Google Sheets')
    else:
        flash(f'Successfully {action.replace("_", " ")}d and synced to Google Sheets')

    return redirect(url_for('student_dashboard'))


# Quiz Routes
@app.route('/admin/quizzes')
def admin_quizzes():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quizzes = Quiz.query.order_by(Quiz.created_at.desc()).all()
    return render_template('admin_quizzes.html', quizzes=quizzes)


@app.route('/admin/quizzes/create', methods=['GET', 'POST'])
def create_quiz():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        title = request.form['title']
        description = request.form['description']
        due_date_str = request.form.get('due_date')

        due_date = None
        if due_date_str:
            due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')

        quiz = Quiz(
            title=title,
            description=description,
            created_by=session['username'],
            due_date=due_date
        )

        db.session.add(quiz)
        db.session.commit()

        flash('Quiz created successfully!')
        return redirect(url_for('admin_quizzes'))

    return render_template('create_quiz.html')


@app.route('/admin/quizzes/<int:quiz_id>/assign', methods=['GET', 'POST'])
def assign_quiz(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)
    students = Student.query.all()

    if request.method == 'POST':
        selected_students = request.form.getlist('students')
        assign_to_all = request.form.get('assign_to_all')

        if assign_to_all == 'on':
            # Assign to all students
            for student in students:
                existing_assignment = QuizAssignment.query.filter_by(
                    quiz_id=quiz_id,
                    student_id=student.student_id
                ).first()

                if not existing_assignment:
                    assignment = QuizAssignment(
                        quiz_id=quiz_id,
                        student_id=student.student_id
                    )
                    db.session.add(assignment)
        else:
            # Assign to selected students only
            for student_id in selected_students:
                existing_assignment = QuizAssignment.query.filter_by(
                    quiz_id=quiz_id,
                    student_id=student_id
                ).first()

                if not existing_assignment:
                    assignment = QuizAssignment(
                        quiz_id=quiz_id,
                        student_id=student_id
                    )
                    db.session.add(assignment)

        db.session.commit()
        flash('Quiz assigned successfully!')
        return redirect(url_for('admin_quizzes'))

    return render_template('assign_quiz.html', quiz=quiz, students=students)


@app.route('/admin/quizzes/<int:quiz_id>/edit', methods=['GET', 'POST'])
def edit_quiz(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)

    if request.method == 'POST':
        quiz.title = request.form['title']
        quiz.description = request.form['description']
        due_date_str = request.form.get('due_date')

        if due_date_str:
            quiz.due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')
        else:
            quiz.due_date = None

        db.session.commit()
        flash('Quiz updated successfully!')
        return redirect(url_for('admin_quizzes'))

    return render_template('edit_quiz.html', quiz=quiz)


@app.route('/admin/quizzes/<int:quiz_id>/delete')
def delete_quiz(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)

    # Delete all assignments related to this quiz
    QuizAssignment.query.filter_by(quiz_id=quiz_id).delete()

    # Delete all questions and options for this quiz
    for question in quiz.questions:
        for option in question.options:
            db.session.delete(option)
        db.session.delete(question)

    db.session.delete(quiz)
    db.session.commit()

    flash('Quiz deleted successfully!')
    return redirect(url_for('admin_quizzes'))


@app.route('/admin/quizzes/<int:quiz_id>/report')
def quiz_report(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)

    # Get all submissions for this quiz
    submissions = QuizSubmission.query.filter_by(quiz_id=quiz_id).all()

    # Calculate statistics
    total_submissions = len(submissions)
    if total_submissions > 0:
        average_score = sum([sub.score for sub in submissions]) / total_submissions
    else:
        average_score = 0

    # Get detailed submission data
    submission_details = []
    for submission in submissions:
        student = Student.query.filter_by(student_id=submission.student_id).first()
        submission_details.append({
            'student': student,
            'submission': submission,
            'percentage': (submission.score / len(quiz.questions)) * 100 if len(quiz.questions) > 0 else 0
        })

    return render_template('quiz_report.html',
                          quiz=quiz,
                          submissions=submission_details,
                          total_submissions=total_submissions,
                          average_score=average_score)


@app.route('/admin/quizzes/<int:quiz_id>/manage_questions')
def manage_quiz_questions(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)
    questions = QuizQuestion.query.filter_by(quiz_id=quiz_id).order_by(QuizQuestion.question_number).all()

    return render_template('manage_quiz_questions.html', quiz=quiz, questions=questions)


@app.route('/admin/quizzes/<int:quiz_id>/add_question', methods=['GET', 'POST'])
def add_quiz_question(quiz_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)

    if request.method == 'POST':
        question_text = request.form['question_text']

        # Get the next question number
        next_question_num = db.session.query(db.func.max(QuizQuestion.question_number)).filter_by(quiz_id=quiz_id).scalar()
        if next_question_num is None:
            next_question_num = 1
        else:
            next_question_num += 1

        # Create the question
        question = QuizQuestion(
            quiz_id=quiz_id,
            question_text=question_text,
            question_number=next_question_num
        )
        db.session.add(question)
        db.session.flush()  # Get the ID for the question

        # Add options
        for i in range(1, 5):  # 4 options
            option_text = request.form.get(f'option_{i}', '')
            is_correct = request.form.get(f'correct_option') == str(i)
            if option_text.strip():  # Only add if option text is not empty
                option = QuizOption(
                    question_id=question.id,
                    option_text=option_text,
                    is_correct=is_correct
                )
                db.session.add(option)

        db.session.commit()
        flash('Question added successfully!')
        return redirect(url_for('manage_quiz_questions', quiz_id=quiz_id))

    return render_template('add_quiz_question.html', quiz=quiz)


@app.route('/admin/quizzes/<int:quiz_id>/edit_question/<int:question_id>', methods=['GET', 'POST'])
def edit_quiz_question(quiz_id, question_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    quiz = Quiz.query.get_or_404(quiz_id)
    question = QuizQuestion.query.filter_by(id=question_id, quiz_id=quiz_id).first_or_404()
    options = QuizOption.query.filter_by(question_id=question_id).all()

    if request.method == 'POST':
        question.question_text = request.form['question_text']

        # Update options
        for i, option in enumerate(options, 1):
            option.option_text = request.form.get(f'option_{i}', '')
            option.is_correct = request.form.get(f'correct_option') == str(i)

        db.session.commit()
        flash('Question updated successfully!')
        return redirect(url_for('manage_quiz_questions', quiz_id=quiz_id))

    return render_template('edit_quiz_question.html', quiz=quiz, question=question, options=options)


@app.route('/admin/quizzes/<int:quiz_id>/delete_question/<int:question_id>')
def delete_quiz_question(quiz_id, question_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    question = QuizQuestion.query.filter_by(id=question_id, quiz_id=quiz_id).first_or_404()

    # Delete all options for this question
    for option in question.options:
        db.session.delete(option)

    db.session.delete(question)
    db.session.commit()

    flash('Question deleted successfully!')
    return redirect(url_for('manage_quiz_questions', quiz_id=quiz_id))


# Assignment Routes
@app.route('/admin/assignments')
def admin_assignments():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    assignments = Assignment.query.order_by(Assignment.created_at.desc()).all()
    return render_template('admin_assignments.html', assignments=assignments)


@app.route('/admin/assignments/create', methods=['GET', 'POST'])
def create_assignment():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        title = request.form['title']
        description = request.form['description']
        due_date_str = request.form.get('due_date')
        file_url = request.form.get('file_url', '')  # Get file URL from form

        due_date = None
        if due_date_str:
            due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')

        assignment = Assignment(
            title=title,
            description=description,
            file_url=file_url,  # Add file URL to assignment
            created_by=session['username'],
            due_date=due_date
        )

        db.session.add(assignment)
        db.session.commit()

        flash('Assignment created successfully!')
        return redirect(url_for('admin_assignments'))

    return render_template('create_assignment.html')


@app.route('/admin/assignments/<int:assignment_id>/assign', methods=['GET', 'POST'])
def assign_assignment(assignment_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    assignment = Assignment.query.get_or_404(assignment_id)
    students = Student.query.all()

    if request.method == 'POST':
        selected_students = request.form.getlist('students')
        assign_to_all = request.form.get('assign_to_all')

        if assign_to_all == 'on':
            # Assign to all students
            for student in students:
                existing_assignment = AssignmentSubmission.query.filter_by(
                    assignment_id=assignment_id,
                    student_id=student.student_id
                ).first()

                if not existing_assignment:
                    assignment_submission = AssignmentSubmission(
                        assignment_id=assignment_id,
                        student_id=student.student_id
                    )
                    db.session.add(assignment_submission)
        else:
            # Assign to selected students only
            for student_id in selected_students:
                existing_assignment = AssignmentSubmission.query.filter_by(
                    assignment_id=assignment_id,
                    student_id=student_id
                ).first()

                if not existing_assignment:
                    assignment_submission = AssignmentSubmission(
                        assignment_id=assignment_id,
                        student_id=student_id
                    )
                    db.session.add(assignment_submission)

        db.session.commit()
        flash('Assignment assigned successfully!')
        return redirect(url_for('admin_assignments'))

    return render_template('assign_assignment.html', assignment=assignment, students=students)


@app.route('/admin/assignments/<int:assignment_id>/edit', methods=['GET', 'POST'])
def edit_assignment(assignment_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    assignment = Assignment.query.get_or_404(assignment_id)

    if request.method == 'POST':
        assignment.title = request.form['title']
        assignment.description = request.form['description']
        assignment.file_url = request.form.get('file_url', '')  # Update file URL
        due_date_str = request.form.get('due_date')

        if due_date_str:
            assignment.due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')
        else:
            assignment.due_date = None

        db.session.commit()
        flash('Assignment updated successfully!')
        return redirect(url_for('admin_assignments'))

    return render_template('edit_assignment.html', assignment=assignment)


@app.route('/admin/assignments/<int:assignment_id>/delete')
def delete_assignment(assignment_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    assignment = Assignment.query.get_or_404(assignment_id)

    # Delete all submissions related to this assignment
    AssignmentSubmission.query.filter_by(assignment_id=assignment_id).delete()

    db.session.delete(assignment)
    db.session.commit()

    flash('Assignment deleted successfully!')
    return redirect(url_for('admin_assignments'))


# Mid-Term Routes
@app.route('/admin/midterms')
def admin_midterms():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    midterms = MidTerm.query.order_by(MidTerm.created_at.desc()).all()
    return render_template('admin_midterms.html', midterms=midterms)


@app.route('/admin/midterms/create', methods=['GET', 'POST'])
def create_midterm():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        title = request.form['title']
        description = request.form['description']
        due_date_str = request.form.get('due_date')
        total_sheets = int(request.form['total_sheets'])
        sheets_per_student = int(request.form.get('sheets_per_student', 5))

        # Handle file upload
        file_url = ""
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename != '':
                # Save file to uploads directory
                import os
                from werkzeug.utils import secure_filename
                filename = secure_filename(file.filename)
                upload_folder = os.path.join(app.root_path, 'static', 'uploads')
                os.makedirs(upload_folder, exist_ok=True)
                file_path = os.path.join(upload_folder, filename)
                file.save(file_path)
                file_url = f"/static/uploads/{filename}"

        due_date = None
        if due_date_str:
            due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')

        midterm = MidTerm(
            title=title,
            description=description,
            file_url=file_url,
            total_sheets=total_sheets,
            sheets_per_student=sheets_per_student,
            created_by=session['username'],
            due_date=due_date
        )

        db.session.add(midterm)
        db.session.commit()

        flash('Mid-term created successfully!')
        return redirect(url_for('admin_midterms'))

    return render_template('create_midterm.html')


@app.route('/admin/midterms/<int:midterm_id>/assign', methods=['GET', 'POST'])
def assign_midterm(midterm_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    midterm = MidTerm.query.get_or_404(midterm_id)
    students = Student.query.all()

    if request.method == 'POST':
        selected_students = request.form.getlist('students')
        assign_to_all = request.form.get('assign_to_all')

        import random

        if assign_to_all == 'on':
            # Assign to all students
            for student in students:
                # Generate random sheet numbers for this student
                available_sheets = list(range(1, midterm.total_sheets + 1))
                assigned_sheets = random.sample(available_sheets, min(midterm.sheets_per_student, len(available_sheets)))

                existing_assignment = MidTermAssignment.query.filter_by(
                    mid_term_id=midterm_id,
                    student_id=student.student_id
                ).first()

                if not existing_assignment:
                    assignment = MidTermAssignment(
                        mid_term_id=midterm_id,
                        student_id=student.student_id,
                        assigned_sheets=','.join(map(str, sorted(assigned_sheets)))
                    )
                    db.session.add(assignment)
        else:
            # Assign to selected students only
            for student_id in selected_students:
                # Generate random sheet numbers for this student
                available_sheets = list(range(1, midterm.total_sheets + 1))
                assigned_sheets = random.sample(available_sheets, min(midterm.sheets_per_student, len(available_sheets)))

                existing_assignment = MidTermAssignment.query.filter_by(
                    mid_term_id=midterm_id,
                    student_id=student_id
                ).first()

                if not existing_assignment:
                    assignment = MidTermAssignment(
                        mid_term_id=midterm_id,
                        student_id=student_id,
                        assigned_sheets=','.join(map(str, sorted(assigned_sheets)))
                    )
                    db.session.add(assignment)

        db.session.commit()
        flash('Mid-term assigned successfully!')
        return redirect(url_for('admin_midterms'))

    return render_template('assign_midterm.html', midterm=midterm, students=students)


@app.route('/admin/midterms/<int:midterm_id>/edit', methods=['GET', 'POST'])
def edit_midterm(midterm_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    midterm = MidTerm.query.get_or_404(midterm_id)

    if request.method == 'POST':
        midterm.title = request.form['title']
        midterm.description = request.form['description']
        due_date_str = request.form.get('due_date')
        midterm.total_sheets = int(request.form['total_sheets'])
        midterm.sheets_per_student = int(request.form.get('sheets_per_student', 5))

        # Handle file upload if provided
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename != '':
                # Save file to uploads directory
                import os
                from werkzeug.utils import secure_filename
                filename = secure_filename(file.filename)
                upload_folder = os.path.join(app.root_path, 'static', 'uploads')
                os.makedirs(upload_folder, exist_ok=True)
                file_path = os.path.join(upload_folder, filename)
                file.save(file_path)
                midterm.file_url = f"/static/uploads/{filename}"

        if due_date_str:
            midterm.due_date = datetime.strptime(due_date_str, '%Y-%m-%dT%H:%M') if 'T' in due_date_str else datetime.strptime(due_date_str, '%Y-%m-%d')
        else:
            midterm.due_date = None

        db.session.commit()
        flash('Mid-term updated successfully!')
        return redirect(url_for('admin_midterms'))

    return render_template('edit_midterm.html', midterm=midterm)


@app.route('/admin/midterms/<int:midterm_id>/delete')
def delete_midterm(midterm_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    midterm = MidTerm.query.get_or_404(midterm_id)

    # Delete all assignments related to this mid-term
    MidTermAssignment.query.filter_by(mid_term_id=midterm_id).delete()

    db.session.delete(midterm)
    db.session.commit()

    flash('Mid-term deleted successfully!')
    return redirect(url_for('admin_midterms'))


@app.route('/admin/midterms/<int:midterm_id>/submissions')
def midterm_submissions(midterm_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    midterm = MidTerm.query.get_or_404(midterm_id)

    # Get all submissions for this mid-term
    submissions = MidTermAssignment.query.filter_by(
        mid_term_id=midterm_id
    ).all()

    return render_template('midterm_submissions.html',
                           midterm=midterm,
                           submissions=submissions)


@app.route('/admin/midterms/<int:midterm_id>/grade/<string:student_id>', methods=['GET', 'POST'])
def grade_midterm_submission(midterm_id, student_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    submission = MidTermAssignment.query.filter_by(
        mid_term_id=midterm_id,
        student_id=student_id
    ).first_or_404()

    if request.method == 'POST':
        grade = float(request.form['grade'])
        submission.grade = grade
        submission.status = 'graded'
        db.session.commit()

        # Add the grade to Google Sheets
        success = add_midterm_grade_to_sheet(
            student_id=submission.student.student_id,
            name=submission.student.name,
            midterm_title=submission.mid_term.title,
            grade=grade,
            graded_at=datetime.now()
        )

        if success:
            flash(f'Mid-term submission for {submission.student.name} graded successfully and recorded in Google Sheets!')
        else:
            flash(f'Mid-term submission for {submission.student.name} graded successfully, but failed to record in Google Sheets.')

        return redirect(url_for('midterm_submissions', midterm_id=midterm_id))

    return render_template('grade_midterm.html', submission=submission)


@app.route('/student/midterms')
def student_midterms():
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']

    # Get mid-terms assigned to this student
    assigned_midterms = db.session.query(MidTerm, MidTermAssignment).join(
        MidTermAssignment
    ).filter(
        MidTermAssignment.student_id == student_id
    ).all()

    return render_template('student_midterms.html', assigned_midterms=assigned_midterms)


@app.route('/student/midterms/<int:midterm_id>/download')
def download_midterm_workbook(midterm_id):
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']
    assignment = MidTermAssignment.query.filter_by(
        mid_term_id=midterm_id,
        student_id=student_id
    ).first()

    if not assignment:
        flash('This mid-term is not assigned to you.', 'error')
        return redirect(url_for('student_midterms'))

    # Here you would generate a custom workbook with only the assigned sheets
    # For now, we'll just return the original file
    import os
    from flask import send_file

    # Get the original file path from the mid-term
    midterm = MidTerm.query.get_or_404(midterm_id)

    # Extract the filename from the stored URL
    if midterm.file_url:
        file_path = os.path.join(app.root_path, 'static', 'uploads', os.path.basename(midterm.file_url))
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)

    flash('Workbook file not found.', 'error')
    return redirect(url_for('student_midterms'))


@app.route('/student/midterms/<int:midterm_id>/submit', methods=['GET', 'POST'])
def submit_midterm(midterm_id):
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']
    midterm = MidTerm.query.get_or_404(midterm_id)

    # Get the assignment record
    assignment = MidTermAssignment.query.filter_by(
        mid_term_id=midterm_id,
        student_id=student_id
    ).first()

    if not assignment:
        flash('This mid-term is not assigned to you.', 'error')
        return redirect(url_for('student_midterms'))

    if request.method == 'POST':
        if 'submission_file' not in request.files:
            flash('No file selected for submission.', 'error')
            return redirect(request.url)

        file = request.files['submission_file']
        if file.filename == '':
            flash('No file selected for submission.', 'error')
            return redirect(request.url)

        if file:
            import os
            from werkzeug.utils import secure_filename
            filename = secure_filename(file.filename)
            upload_folder = os.path.join(app.root_path, 'static', 'submissions')
            os.makedirs(upload_folder, exist_ok=True)

            # Create a unique filename using student_id and midterm_id
            unique_filename = f"midterm_{midterm_id}_student_{student_id}_{filename}"
            file_path = os.path.join(upload_folder, unique_filename)
            file.save(file_path)

            # Update the assignment record with the submission file path
            assignment.submission_file_path = f"/static/submissions/{unique_filename}"
            assignment.status = 'submitted'
            assignment.submitted_at = datetime.now()
            db.session.commit()

            flash('Mid-term submitted successfully!', 'success')
            return redirect(url_for('student_midterms'))

    return render_template('submit_midterm.html', midterm=midterm, assignment=assignment)


@app.route('/student/quizzes')
def student_quizzes():
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']

    # Get quizzes assigned to this student
    assigned_quizzes = db.session.query(Quiz, QuizAssignment).join(
        QuizAssignment
    ).filter(
        QuizAssignment.student_id == student_id
    ).all()

    return render_template('student_quizzes.html', assigned_quizzes=assigned_quizzes)


@app.route('/student/quizzes/<int:quiz_id>/take', methods=['GET', 'POST'])
def take_quiz(quiz_id):
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']
    quiz = Quiz.query.get_or_404(quiz_id)

    # Check if quiz is assigned to this student
    quiz_assignment = QuizAssignment.query.filter_by(
        quiz_id=quiz_id,
        student_id=student_id
    ).first()

    if not quiz_assignment:
        flash('This quiz is not assigned to you.', 'error')
        return redirect(url_for('student_quizzes'))

    # Check if quiz is already submitted
    existing_submission = QuizSubmission.query.filter_by(
        quiz_id=quiz_id,
        student_id=student_id
    ).first()

    if existing_submission:
        flash('You have already taken this quiz.', 'info')
        return redirect(url_for('view_quiz_result', quiz_id=quiz_id))

    questions = QuizQuestion.query.filter_by(quiz_id=quiz_id).order_by(QuizQuestion.question_number).all()

    if request.method == 'POST':
        # Process the quiz submission
        correct_answers = 0
        total_questions = len(questions)

        # Create a new submission
        submission = QuizSubmission(
            quiz_id=quiz_id,
            student_id=student_id,
            score=0  # Will be calculated after processing answers
        )
        db.session.add(submission)
        db.session.flush()  # Get the submission ID

        # Process each question
        for question in questions:
            selected_option_id = request.form.get(f'question_{question.id}')

            if selected_option_id:
                # Find the selected option
                selected_option = QuizOption.query.get(selected_option_id)

                # Determine if the answer is correct
                is_correct_answer = selected_option.is_correct if selected_option else False

                # Create the answer record
                answer = QuizAnswer(
                    submission_id=submission.id,
                    question_id=question.id,
                    selected_option_id=selected_option_id,
                    is_correct=is_correct_answer
                )
                db.session.add(answer)

                # Increment correct answers if the answer is correct
                if is_correct_answer:
                    correct_answers += 1

        # Update the submission score
        submission.score = correct_answers
        db.session.commit()

        # Update the quiz assignment status and grade
        quiz_assignment.status = 'graded'  # Since scoring is automatic, mark as graded
        quiz_assignment.grade = (correct_answers / total_questions) * 100 if total_questions > 0 else 0
        db.session.commit()

        # Get student info for Google Sheets
        student = Student.query.filter_by(student_id=student_id).first()

        # Add to Google Sheets
        success = add_quiz_submission_to_sheet(
            student_id=student.student_id,
            name=student.name,
            quiz_title=quiz.title,
            score=correct_answers,
            total_questions=total_questions,
            submitted_at=submission.submitted_at
        )

        # Also save detailed answers to Google Sheets
        # We need to get the questions and answers data for the detailed report
        questions_data = []
        for question in questions:
            # Find the answer for this question in this submission
            answer = next((ans for ans in submission.answers if ans.question_id == question.id), None)

            selected_option = None
            is_correct = False
            if answer:
                selected_option = answer.selected_option
                is_correct = answer.is_correct

            # Get the correct option for this question
            correct_option = QuizOption.query.filter_by(question_id=question.id, is_correct=True).first()

            questions_data.append({
                'question': question,
                'selected_option': selected_option,
                'correct_option': correct_option,
                'is_correct': is_correct
            })

        # Save detailed answers to Google Sheets
        detailed_success = add_detailed_quiz_answers_to_sheet(
            student_id=student.student_id,
            name=student.name,
            quiz_title=quiz.title,
            questions_data=questions_data,
            submitted_at=submission.submitted_at
        )

        if not success or not detailed_success:
            if not success and not detailed_success:
                flash('Quiz submitted locally but failed to sync to Google Sheets', 'warning')
            elif not success:
                flash('Quiz submitted but summary failed to sync to Google Sheets', 'warning')
            else:
                flash('Quiz submitted but detailed answers failed to sync to Google Sheets', 'warning')
        else:
            flash(f'Quiz submitted successfully! You scored {correct_answers}/{total_questions}.', 'success')

        return redirect(url_for('view_quiz_result', quiz_id=quiz_id))

    return render_template('take_quiz.html', quiz=quiz, questions=questions)


@app.route('/student/quizzes/<int:quiz_id>/result')
def view_quiz_result(quiz_id):
    if 'student_id' not in session and 'admin_id' not in session:
        return redirect(url_for('login'))

    # Determine if accessed by student or admin
    if 'student_id' in session:
        student_id = session['student_id']
    else:
        # Admin accessing - check if student_id is provided in query params
        student_id = request.args.get('student_id')
        if not student_id:
            flash('Student ID not provided.', 'error')
            return redirect(url_for('admin_quizzes'))

    quiz = Quiz.query.get_or_404(quiz_id)

    # Get the submission for this quiz
    submission = QuizSubmission.query.filter_by(
        quiz_id=quiz_id,
        student_id=student_id
    ).first()

    if not submission:
        flash('Quiz submission not found.', 'error')
        if 'admin_id' in session:
            return redirect(url_for('quiz_report', quiz_id=quiz_id))
        else:
            return redirect(url_for('student_quizzes'))

    # Get questions and answers
    # Get all questions for this quiz
    questions = QuizQuestion.query.filter_by(quiz_id=quiz.id).order_by(QuizQuestion.question_number).all()

    questions_data = []
    for question in questions:
        # Find the answer for this question in this submission
        answer = next((ans for ans in submission.answers if ans.question_id == question.id), None)

        selected_option = None
        is_correct = False
        if answer:
            selected_option = answer.selected_option
            is_correct = answer.is_correct

        # Get the correct option for this question
        correct_option = QuizOption.query.filter_by(question_id=question.id, is_correct=True).first()

        questions_data.append({
            'question': question,
            'selected_option': selected_option,
            'correct_option': correct_option,
            'is_correct': is_correct
        })

    return render_template('quiz_result.html', quiz=quiz, submission=submission, questions_data=questions_data)


@app.route('/student/assignments')
def student_assignments():
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']

    # Get assignments assigned to this student
    assigned_assignments = db.session.query(Assignment, AssignmentSubmission).join(
        AssignmentSubmission
    ).filter(
        AssignmentSubmission.student_id == student_id
    ).all()

    return render_template('student_assignments.html', assigned_assignments=assigned_assignments)


@app.route('/student/assignments/<int:assignment_id>/submit', methods=['GET', 'POST'])
def submit_assignment(assignment_id):
    if 'student_id' not in session:
        return redirect(url_for('login'))

    student_id = session['student_id']
    assignment = Assignment.query.get_or_404(assignment_id)

    # Get the assignment submission record
    submission = AssignmentSubmission.query.filter_by(
        assignment_id=assignment_id,
        student_id=student_id
    ).first()

    if not submission:
        flash('This assignment is not assigned to you.', 'error')
        return redirect(url_for('student_assignments'))

    if request.method == 'POST':
        submission_url = request.form.get('submission_url', '')

        # Update submission status and URL
        submission.submission_url = submission_url
        submission.status = 'submitted'
        submission.submitted_at = datetime.now()

        db.session.commit()

        # Get student info for Google Sheets
        student = Student.query.filter_by(student_id=student_id).first()

        # Add to Google Sheets
        success = add_assignment_submission_to_sheet(
            student_id=student.student_id,
            name=student.name,
            assignment_title=assignment.title,
            submission_url=submission_url,
            submitted_at=submission.submitted_at
        )

        if not success:
            flash('Assignment submitted locally but failed to sync to Google Sheets', 'warning')
        else:
            flash('Assignment submitted successfully and synced to Google Sheets!', 'success')

        return redirect(url_for('student_assignments'))

    return render_template('submit_assignment.html', assignment=assignment, submission=submission)


@app.route('/admin/course-outlines')
def admin_course_outlines():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    course_outlines = CourseOutline.query.order_by(CourseOutline.week_number, CourseOutline.created_at.desc()).all()
    return render_template('admin_course_outlines.html', course_outlines=course_outlines)


@app.route('/admin/course-outlines/create', methods=['GET', 'POST'])
def create_course_outline():
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        title = request.form['title']
        description = request.form['description']
        week_number = request.form.get('week_number', type=int)
        outline_file_url = request.form.get('outline_file_url', '')
        presentation_file_url = request.form.get('presentation_file_url', '')

        course_outline = CourseOutline(
            title=title,
            description=description,
            outline_file_url=outline_file_url,
            presentation_file_url=presentation_file_url,
            week_number=week_number,
            created_by=session['username']
        )

        db.session.add(course_outline)
        db.session.commit()

        flash('Course outline created successfully!')
        return redirect(url_for('admin_course_outlines'))

    return render_template('create_course_outline.html')


@app.route('/admin/course-outlines/<int:outline_id>/edit', methods=['GET', 'POST'])
def edit_course_outline(outline_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    course_outline = CourseOutline.query.get_or_404(outline_id)

    if request.method == 'POST':
        course_outline.title = request.form['title']
        course_outline.description = request.form['description']
        course_outline.week_number = request.form.get('week_number', type=int)
        course_outline.outline_file_url = request.form.get('outline_file_url', '')
        course_outline.presentation_file_url = request.form.get('presentation_file_url', '')
        course_outline.updated_at = datetime.now()

        db.session.commit()

        flash('Course outline updated successfully!')
        return redirect(url_for('admin_course_outlines'))

    return render_template('edit_course_outline.html', course_outline=course_outline)


@app.route('/admin/course-outlines/<int:outline_id>/delete')
def delete_course_outline(outline_id):
    if 'admin_id' not in session:
        return redirect(url_for('login'))

    course_outline = CourseOutline.query.get_or_404(outline_id)
    db.session.delete(course_outline)
    db.session.commit()

    flash('Course outline deleted successfully!')
    return redirect(url_for('admin_course_outlines'))


@app.route('/student/course-outlines')
def student_course_outlines():
    if 'student_id' not in session:
        return redirect(url_for('login'))

    # Get all active course outlines ordered by week number
    course_outlines = CourseOutline.query.filter_by(is_active=True).order_by(CourseOutline.week_number).all()
    return render_template('student_course_outlines.html', course_outlines=course_outlines)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


if __name__ == '__main__':
    with app.app_context():
        db.create_all()

        # Create default admin if not exists
        admin = Admin.query.filter_by(username='admin').first()
        if not admin:
            admin = Admin(username='admin')
            admin.set_password('admin123')
            db.session.add(admin)
            db.session.commit()

    app.run(debug=True)