# Academic ERP System

An Enterprise Resource Planning system built with Flask, featuring student management, attendance tracking, quizzes, assignments, and mid-term examinations.

## Features

- Admin dashboard for managing students and courses
- Student portal for accessing quizzes, assignments, and course materials
- Attendance tracking with Google Sheets integration
- Mid-term examination system with Excel workbook distribution
- Course outline management with presentation uploads
- Professional dark-themed UI/UX

## Tech Stack

- Python Flask
- SQLite Database
- Bootstrap 5
- Google Sheets API
- Google Drive API

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/muaazasif/academic-erp-system.git
   cd academic-erp-system
   ```

2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up environment variables:
   - Create a `.env` file with your Google Sheets credentials
   - Add your Google Sheet ID

5. Run the application:
   ```bash
   python app.py
   ```

## Environment Variables

Create a `.env` file with the following variables:
- `GOOGLE_SHEET_ID` - Your Google Sheet ID
- `GOOGLE_SHEETS_CREDENTIALS_JSON` - Your Google Sheets service account credentials in JSON format
- `SECRET_KEY` - Your Flask secret key

## Admin Credentials

- Username: `admin`
- Password: `admin123`

## Deployment Options

### For Production Deployment:

#### Option 1: Google Cloud Run (Recommended for Google Services)
Students can access the application from home!

How?
Deploy the app on Google Cloud Run

After deployment, you'll get a URL like:
```
https://myerp-app-xyz.a.run.app
```

**Benefits:**
- ✔️ 100+ users supported easily
- ✔️ Auto-scaling (handles traffic spikes automatically)
- ✔️ Mostly free under Google Cloud free tier
- ✔️ Students can access from anywhere (home, college, etc.)
- ✔️ High performance and reliability
- ✔️ Built-in security and HTTPS

#### Option 2: PythonAnywhere.com (Recommended for Python Apps!)
Students can access the application from home too!

How?
Deploy the app on PythonAnywhere.com

After deployment, you'll get a URL like:
```
https://your-username.pythonanywhere.com
```

**Benefits:**
- ✔️ 100+ users supported easily
- ✔️ Auto-scaling (based on plan)
- ✔️ Free tier available with paid upgrade options
- ✔️ Students can access from anywhere (home, college, etc.)
- ✔️ No server management required
- ✔️ Python-optimized environment
- ✔️ SSL certificate available
- ✔️ Integrated with Google services

#### Option 3: Railway.app (Modern Cloud Platform!)
Students can access the application from home too!

How?
Deploy the app on Railway.app

After deployment, you'll get a URL like:
```
https://your-project-name.up.railway.app
```

**Benefits:**
- ✔️ 100+ users supported easily
- ✔️ Auto-scaling (when properly configured)
- ✔️ Free tier available with paid upgrade options
- ✔️ Students can access from anywhere (home, college, etc.)
- ✔️ No server management required
- ✔️ Modern deployment platform
- ✔️ SSL certificate automatically provided
- ✔️ Docker-based deployment

## License

This project is licensed under the MIT License - see the LICENSE file for details.