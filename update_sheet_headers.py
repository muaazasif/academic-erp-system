#!/usr/bin/env python3
"""
Script to update Google Sheets with proper headers for location-based attendance tracking
"""

import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service object"""
    creds = None

    # Load the full Google service account JSON from environment variable
    creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    if creds_json:
        try:
            # Parse the credentials JSON from environment variable
            info = json.loads(creds_json)

            # Restore newlines in the private key to fix PEM formatting
            if "private_key" in info:
                private_key = info["private_key"]

                # Handle various newline representations that might be in the environment variable
                # Replace escaped newlines (common when storing in env vars)
                private_key = private_key.replace("\\n", "\n").replace("\\r", "\r")

                # Ensure proper PEM formatting
                if not private_key.startswith("-----BEGIN"):
                    private_key = "-----BEGIN PRIVATE KEY-----\n" + private_key
                if not private_key.endswith("-----END PRIVATE KEY-----\n"):
                    if not private_key.endswith("\n"):
                        private_key += "\n"
                    private_key += "-----END PRIVATE KEY-----\n"

                info["private_key"] = private_key

            # Create credentials object from the service account info
            creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets'])
            print("Successfully loaded credentials from environment variable GOOGLE_SHEETS_CREDENTIALS_JSON")
        except (ValueError, TypeError) as env_error:
            print(f"Error parsing service account credentials from environment: {env_error}")
            import traceback
            traceback.print_exc()
            return None
        except Exception as general_error:
            print(f"General error with environment credentials: {general_error}")
            import traceback
            traceback.print_exc()
            return None

    # If we have valid credentials from environment variable, use them; otherwise, return None gracefully
    if not creds:
        print("No valid credentials found from environment variable GOOGLE_SHEETS_CREDENTIALS_JSON. Google Sheets integration will be disabled.")
        return None

    try:
        service = build('sheets', 'v4', credentials=creds)
        return service
    except Exception as e:
        print(f"Error building Google Sheets service: {e}")
        import traceback
        traceback.print_exc()
        return None

def update_attendance_sheet_headers():
    """Update the attendance sheet with proper headers for location data"""
    service = authenticate_google_sheets()
    
    if not service:
        print("Cannot authenticate to Google Sheets. Please check your credentials.")
        return False
    
    # Get the Google Sheet ID from environment variable
    SPREADSHEET_ID = os.getenv('GOOGLE_SHEET_ID')
    
    if not SPREADSHEET_ID:
        print("Google Sheet ID not found in environment variables. Please set GOOGLE_SHEET_ID.")
        return False
    
    try:
        sheet = service.spreadsheets()
        
        # Define the headers for the attendance sheet with location data
        headers = [
            'Date', 'Student ID', 'Name', 'Check-In Time', 'Check-Out Time', 'Status',
            'Check-In Location (Lat,Lng)', 'Check-In Latitude', 'Check-In Longitude', 'Check-In Address',
            'Check-Out Location (Lat,Lng)', 'Check-Out Latitude', 'Check-Out Longitude', 'Check-Out Address',
            'Sync Timestamp'
        ]
        
        # Prepare the header data
        header_values = [headers]
        
        body = {
            'values': header_values
        }
        
        # Try to update the Attendance sheet
        try:
            # First, get the spreadsheet to see what sheets are available
            spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            available_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
            
            print(f"Available sheets in spreadsheet: {available_sheets}")
            
            # Check if 'Attendance' sheet exists
            target_sheet = 'Attendance'
            if target_sheet not in available_sheets:
                # If Attendance sheet doesn't exist, try other common names
                for sheet_name in ['Sheet1', 'Data', 'Records']:
                    if sheet_name in available_sheets:
                        target_sheet = sheet_name
                        print(f"Using '{target_sheet}' sheet since 'Attendance' doesn't exist")
                        break
            
            # Update the header row in the target sheet
            range_name = f'{target_sheet}!A1:P1'  # A1 to P1 for 16 columns
            
            result = sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            # Format the header row to make it stand out
            format_request = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": None,  # Will be filled in below
                                "startRowIndex": 0,
                                "endRowIndex": 1
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": 0.2,
                                        "green": 0.2,
                                        "blue": 0.8  # Blue background
                                    },
                                    "textFormat": {
                                        "bold": True,
                                        "foregroundColor": {
                                            "red": 1.0,
                                            "green": 1.0,
                                            "blue": 1.0  # White text
                                        }
                                    }
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,textFormat)"
                        }
                    }
                ]
            }
            
            # Get the sheet ID for the target sheet to apply formatting
            spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            for sheet_info in spreadsheet['sheets']:
                if sheet_info['properties']['title'] == target_sheet:
                    format_request['requests'][0]['repeatCell']['range']['sheetId'] = sheet_info['properties']['sheetId']
                    break
            
            # Apply formatting
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=format_request
            ).execute()
            
            print(f"Headers updated successfully in '{target_sheet}' sheet.")
            print(f"Updated {result.get('updatedCells')} cells with location-based attendance headers.")
            print("New columns added:")
            for i, header in enumerate(headers, 1):
                print(f"  {i:2d}. {header}")
            
            return True
            
        except HttpError as e:
            print(f"HTTP Error: {e}")
            print("This might happen if you don't have write permissions to the sheet.")
            return False
        except Exception as e:
            print(f"Error updating attendance sheet headers: {e}")
            import traceback
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Updating Google Sheets with location-based attendance headers...")
    success = update_attendance_sheet_headers()
    if success:
        print("\n✓ Google Sheet headers updated successfully!")
        print("You should now see location data in your Google Sheet when new attendance records are added.")
    else:
        print("\n✗ Google Sheet headers update failed!")
        print("Make sure your Google Sheets credentials are correct and you have write permissions.")