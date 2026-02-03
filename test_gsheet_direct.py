import os
import sys
import json
from dotenv import load_dotenv
from datetime import datetime

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Load environment variables
load_dotenv()

# Import the app module to access its functions
from app import authenticate_google_sheets, add_quiz_submission_to_sheet, add_detailed_quiz_answers_to_sheet

def test_quiz_submission():
    """Test sending quiz data to Google Sheets"""
    print("Testing quiz submission to Google Sheets...")
    
    # Test basic quiz submission
    success = add_quiz_submission_to_sheet(
        student_id="101",
        name="Muaaz",
        quiz_title="Test Quiz",
        score=1,
        total_questions=1,
        submitted_at=datetime.now()
    )
    
    print(f"Basic quiz submission result: {success}")
    
    # Test detailed answers submission
    # Create mock questions data
    questions_data = [{
        'question': type('obj', (object,), {'question_number': 1, 'question_text': 'How'})(),
        'selected_option': type('obj', (object,), {'option_text': 's'})(),
        'correct_option': type('obj', (object,), {'option_text': 's'})(),
        'is_correct': True
    }]
    
    detailed_success = add_detailed_quiz_answers_to_sheet(
        student_id="101",
        name="Muaaz",
        quiz_title="Test Quiz",
        questions_data=questions_data,
        submitted_at=datetime.now()
    )
    
    print(f"Detailed answers submission result: {detailed_success}")
    
    return success and detailed_success

if __name__ == "__main__":
    print("Starting Google Sheets integration test...")
    result = test_quiz_submission()
    print(f"Test completed. Success: {result}")