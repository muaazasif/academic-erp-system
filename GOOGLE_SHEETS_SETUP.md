# Google Sheets Integration Setup Guide

To enable attendance data synchronization to Google Sheets, follow these steps:

## Step 1: Create a Google Cloud Project
1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Sheets API for your project

## Step 2: Create OAuth 2.0 Credentials
1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "OAuth 2.0 Client IDs"
3. Choose "Desktop application" as the application type
4. Download the credentials JSON file
5. Rename the downloaded file to `credentials.json` and place it in the `myerp_app` directory

## Step 3: Configure Google Sheets Permissions
1. Create a new Google Sheet or use an existing one
2. Note the Sheet ID from the URL: `https://docs.google.com/spreadsheets/d/[YOUR_SHEET_ID]/edit`
3. Share the Google Sheet with the email address found in your `credentials.json` file (or the email associated with the service account)

## Step 4: Update the Application Configuration
1. Open `app.py` in the `myerp_app` directory
2. Find the `add_attendance_to_sheet` function
3. Replace `'YOUR_SPREADSHEET_ID_HERE'` with your actual Google Sheet ID
4. If your sheet has a different name than 'Sheet1', update the range parameter accordingly

## Step 5: Install Dependencies
Run the following command to install the required Google API libraries:
```
pip install -r requirements.txt
```

## Step 6: Run the Application
Start the application:
```
python app.py
```

The first time you record attendance, you'll be prompted to authenticate with Google. Follow the browser prompts to complete the authentication process. A `token.pickle` file will be created to store your credentials for future use.

## Important Notes:
- The application will store attendance data in both the local database and Google Sheets
- If Google Sheets sync fails, the data will still be stored locally
- Make sure your Google Sheet has the correct permissions for the application to write data
- The data will be appended to the sheet in the following columns: Date, Student ID, Name, Check-In Time, Check-Out Time, Status, Timestamp