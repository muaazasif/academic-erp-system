# Location-Based Attendance Tracking

This ERP system includes location-based attendance tracking functionality that allows students to check in and out with their GPS coordinates.

## Features

- **Check-in/Check-out**: Students can record their arrival and departure times
- **Location Tracking**: Captures GPS coordinates (latitude and longitude) for both check-in and check-out
- **Google Sheets Integration**: Automatically syncs attendance records to Google Sheets with location data
- **Database Storage**: Stores attendance records locally with location information

## How It Works

1. **Student Dashboard**: Students access their dashboard and click "Check In" or "Check Out" buttons
2. **Location Access**: The system requests permission to access the student's location via browser geolocation API
3. **Data Capture**: Records date, time, and GPS coordinates for both check-in and check-out
4. **Database Storage**: Saves the attendance record to the local SQLite database
5. **Google Sheets Sync**: Automatically sends the data to Google Sheets with the following columns:
   - Date
   - Student ID
   - Name
   - Check-In Time
   - Check-Out Time
   - Status
   - Check-In Location (Lat,Lng)
   - Check-In Latitude
   - Check-In Longitude
   - Check-Out Location (Lat,Lng)
   - Check-Out Latitude
   - Check-Out Longitude
   - Sync Timestamp

## Technical Implementation

### Database Schema
The `attendance` table includes these location-related columns:
- `check_in_location`: Stores the full location string (latitude,longitude) for check-in
- `check_out_location`: Stores the full location string (latitude,longitude) for check-out

### Google Sheets Integration
- Enhanced `add_attendance_to_sheet()` function splits location coordinates into separate latitude and longitude columns
- Automatic retry mechanism for failed syncs using `sync_utils.py`
- Proper error handling and logging

### Frontend Implementation
- JavaScript geolocation API integration in `student_dashboard.html`
- Permission handling for location access
- Form submission with location data

## Setup Requirements

1. **Google Sheets API Credentials**: Set up service account credentials and configure environment variables:
   - `GOOGLE_SHEETS_CREDENTIALS_JSON`: Full service account JSON credentials
   - `GOOGLE_SHEET_ID`: Target Google Sheet ID

2. **Database Migration**: Run the migration script to ensure location columns exist

## Privacy Considerations

- Location data is only collected when students actively check in/out
- Students are prompted for location permission by the browser
- Location data is stored securely in encrypted Google Sheets
- Only authorized administrators can access the full location data