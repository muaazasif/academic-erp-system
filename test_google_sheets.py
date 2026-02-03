import os
import sys
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Add the current directory to Python path to import app module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app import authenticate_google_sheets, add_attendance_to_sheet

def test_google_sheets_connection():
    """Test the Google Sheets connection and functionality"""
    print("Testing Google Sheets connection...")

    try:
        # Test authentication
        service = authenticate_google_sheets()
        print("[SUCCESS] Successfully authenticated with Google Sheets")

        # Test adding a sample record
        success = add_attendance_to_sheet(
            student_id="TEST001",
            name="Test Student",
            date="2026-02-02",
            check_in="10:00:00",
            check_out="18:00:00",
            status="present"
        )

        if success:
            print("[SUCCESS] Successfully added test record to Google Sheets")
        else:
            print("[ERROR] Failed to add record to Google Sheets")

    except Exception as e:
        print(f"[ERROR] Error during Google Sheets test: {e}")

if __name__ == "__main__":
    test_google_sheets_connection()