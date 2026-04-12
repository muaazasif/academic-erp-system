"""
Test Google Sheets Connection
Run this to check if your credentials are working
"""
import os
from app import get_google_sheets_service

print("=" * 70)
print("🔍 TESTING GOOGLE SHEETS CONNECTION")
print("=" * 70)

# Check environment variables
print("\n📋 Environment Variables:")
sheet_id = os.getenv('GOOGLE_SHEET_ID')
creds = os.getenv('GOOGLE_SHEETS_CREDENTIALS_JSON')

if sheet_id:
    print(f"✅ GOOGLE_SHEET_ID: {sheet_id}")
else:
    print("❌ GOOGLE_SHEET_ID: NOT SET")

if creds:
    print(f"✅ GOOGLE_SHEETS_CREDENTIALS_JSON: SET ({len(creds)} chars)")
else:
    print("❌ GOOGLE_SHEETS_CREDENTIALS_JSON: NOT SET")

# Try to connect
print("\n🔄 Testing connection...")
service, connected_sheet_id = get_google_sheets_service()

if service:
    print(f"✅ Google Sheets service connected!")
    print(f"✅ Sheet ID: {connected_sheet_id}")
    
    # Try to access the spreadsheet
    try:
        spreadsheet = service.spreadsheets().get(spreadsheetId=connected_sheet_id).execute()
        print(f"✅ Spreadsheet accessible: {spreadsheet.get('properties', {}).get('title', 'Unknown')}")
        
        sheets = spreadsheet.get('sheets', [])
        print(f"📊 Available sheets ({len(sheets)}):")
        for sheet in sheets:
            print(f"   - {sheet['properties']['title']}")
        
        print("\n🎉 ALL TESTS PASSED! Your Google Sheets sync will work!")
    except Exception as e:
        print(f"❌ Error accessing spreadsheet: {e}")
else:
    print("❌ Google Sheets service NOT available")
    print("\n💡 FIX:")
    print("   1. Set GOOGLE_SHEET_ID in Railway variables")
    print("   2. Set GOOGLE_SHEETS_CREDENTIALS_JSON in Railway variables")
    print("   3. Share your Google Sheet with the service account email")

print("=" * 70)
