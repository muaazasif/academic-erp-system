import os
import sys
from datetime import datetime
from dotenv import load_dotenv

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Load environment variables
load_dotenv()

# Import the app module to access its functions
from app import authenticate_google_sheets, add_quiz_submission_to_sheet, add_detailed_quiz_answers_to_sheet

def test_google_sheets_functions():
    """Test the Google Sheets functions directly"""
    print("Testing Google Sheets functions...")
    
    try:
        # Test basic quiz submission
        print("\n1. Testing basic quiz submission...")
        success = add_quiz_submission_to_sheet(
            student_id="101",
            name="Muaaz",
            quiz_title="Test Quiz",
            score=0,  # Testing with 0 score (incorrect answer)
            total_questions=1,
            submitted_at=datetime.now()
        )
        print(f"   Basic quiz submission result: {success}")
        
        if success:
            print("   SUCCESS: Quiz data was successfully sent to Google Sheets!")
        else:
            print("   FAILED: Quiz data failed to send to Google Sheets")

        # Test detailed answers submission
        print("\n2. Testing detailed answers submission...")

        # Create mock questions data for an incorrect answer
        questions_data = [{
            'question': type('obj', (object,), {
                'question_number': 1,
                'question_text': 'How'
            })(),
            'selected_option': type('obj', (object,), {
                'option_text': 's'
            })(),
            'correct_option': type('obj', (object,), {
                'option_text': 's'
            })(),
            'is_correct': False  # This represents an incorrect answer
        }]

        detailed_success = add_detailed_quiz_answers_to_sheet(
            student_id="101",
            name="Muaaz",
            quiz_title="Test Quiz",
            questions_data=questions_data,
            submitted_at=datetime.now()
        )

        print(f"   Detailed answers submission result: {detailed_success}")

        if detailed_success:
            print("   SUCCESS: Detailed answers data was successfully sent to Google Sheets!")
        else:
            print("   FAILED: Detailed answers data failed to send to Google Sheets")

        print(f"\nOverall result: {'SUCCESS' if success and detailed_success else 'FAILED'}")
        return success and detailed_success
        
    except Exception as e:
        print(f"Error during Google Sheets test: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Starting Google Sheets integration test...")
    result = test_google_sheets_functions()
    print(f"\nTest completed. Overall success: {result}")