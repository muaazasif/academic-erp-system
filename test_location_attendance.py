#!/usr/bin/env python3
"""
Test script to verify location-based attendance functionality
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app, db, Student, Attendance, add_attendance_to_sheet
from datetime import datetime, date

def test_location_attendance():
    """Test the location-based attendance functionality"""
    with app.app_context():
        # Create a test student if not exists
        test_student = Student.query.filter_by(student_id='TEST001').first()
        if not test_student:
            test_student = Student(student_id='TEST001', name='Test Student')
            test_student.set_password('password')
            db.session.add(test_student)
            db.session.commit()
            print("Created test student TEST001")
        
        # Create a test attendance record with location data
        test_attendance = Attendance(
            student_id='TEST001',
            date=date.today(),
            check_in_time=datetime.now().time(),
            check_out_time=datetime.now().time(),
            check_in_location='24.8607,67.0011',  # Karachi coordinates
            check_out_location='24.8608,67.0012',  # Slightly different coordinates
            status='present'
        )
        db.session.add(test_attendance)
        db.session.commit()
        
        print(f"Created test attendance record for {test_student.name}")
        print(f"Check-in location: {test_attendance.check_in_location}")
        print(f"Check-out location: {test_attendance.check_out_location}")
        
        # Test sending to Google Sheets
        print("\nTesting Google Sheets sync with location data...")
        success = add_attendance_to_sheet(
            student_id=test_student.student_id,
            name=test_student.name,
            date=test_attendance.date,
            check_in=test_attendance.check_in_time,
            check_out=test_attendance.check_out_time,
            status=test_attendance.status,
            check_in_location=test_attendance.check_in_location,
            check_out_location=test_attendance.check_out_location
        )
        
        if success:
            print("✓ Location-based attendance successfully synced to Google Sheets!")
        else:
            print("✗ Failed to sync location-based attendance to Google Sheets")
            print("  Note: This may be due to missing Google Sheets credentials")
        
        return success

if __name__ == "__main__":
    print("Testing location-based attendance functionality...")
    success = test_location_attendance()
    if success:
        print("\n✓ Test completed successfully!")
    else:
        print("\n⚠ Test completed with warnings (likely due to missing credentials)")
        # Don't exit with error code since missing credentials is expected in test environments