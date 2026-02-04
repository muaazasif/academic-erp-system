#!/usr/bin/env python3
"""
Initialize Google Sheets with proper headers for location-based attendance tracking
"""

import os
import sys
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def authenticate_google_sheets():
    """Authenticate and return Google Sheets service object"""
    creds = None
    from google.oauth2.service_account import Credentials

    # Load the full Google service account JSON from environment variable
    creds_json = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')
    if creds_json:
        import json
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

def initialize_attendance_sheet():
    """Initialize the attendance sheet with proper headers"""
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
        
        # Define the headers for the attendance sheet
        headers = [
            'Date', 'Student ID', 'Name', 'Check-In Time', 'Check-Out Time', 'Status',
            'Check-In Location (Lat,Lng)', 'Check-In Latitude', 'Check-In Longitude',
            'Check-Out Location (Lat,Lng)', 'Check-Out Latitude', 'Check-Out Longitude',
            'Sync Timestamp'
        ]
        
        # Prepare the header data
        header_values = [headers]
        
        body = {
            'values': header_values
        }
        
        # Try to update the Attendance sheet, create it if it doesn't exist
        try:
            # First, get the spreadsheet to see what sheets are available
            spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            available_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
            
            print(f"Available sheets in spreadsheet: {available_sheets}")
            
            # Check if 'Attendance' sheet exists
            if 'Attendance' not in available_sheets:
                # Create a new sheet named 'Attendance'
                add_sheet_request = {
                    "requests": [{
                        "addSheet": {
                            "properties": {
                                "title": "Attendance",
                                "gridProperties": {
                                    "rowCount": 1000,
                                    "columnCount": 20
                                }
                            }
                        }
                    }]
                }
                
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=add_sheet_request
                ).execute()
                
                print("Created new 'Attendance' sheet")
            
            # Clear any existing data in A1:M1 range and add headers
            result = sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range='Attendance!A1:M1',
                valueInputOption='RAW',
                body=body
            ).execute()
            
            # Format the header row
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
                                        "blue": 0.2
                                    },
                                    "textFormat": {
                                        "bold": True,
                                        "foregroundColor": {
                                            "red": 1.0,
                                            "green": 1.0,
                                            "blue": 1.0
                                        }
                                    }
                                }
                            },
                            "fields": "userEnteredFormat(backgroundColor,textFormat)"
                        }
                    }
                ]
            }
            
            # Get the sheet ID for the Attendance sheet to apply formatting
            spreadsheet = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
            for sheet_info in spreadsheet['sheets']:
                if sheet_info['properties']['title'] == 'Attendance':
                    format_request['requests'][0]['repeatCell']['range']['sheetId'] = sheet_info['properties']['sheetId']
                    break
            
            # Apply formatting
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=format_request
            ).execute()
            
            print(f"Headers initialized successfully in 'Attendance' sheet.")
            print(f"Updated {result.get('updatedCells')} cells.")
            
            return True
            
        except HttpError as e:
            print(f"HTTP Error: {e}")
            print("This might happen if the sheet doesn't exist or you don't have write permissions.")
            return False
        except Exception as e:
            print(f"Error initializing attendance sheet: {e}")
            import traceback
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Initializing Google Sheets for location-based attendance tracking...")
    success = initialize_attendance_sheet()
    if success:
        print("Google Sheet initialization completed successfully!")
    else:
        print("Google Sheet initialization failed!")
        sys.exit(1)