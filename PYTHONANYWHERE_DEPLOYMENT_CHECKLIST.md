# PythonAnywhere Deployment Checklist

## Pre-deployment Steps:
- [ ] Create PythonAnywhere account at https://www.pythonanywhere.com/
- [ ] Download all project files to your local machine
- [ ] Ensure you have the following files ready:
  - app.py (main Flask application)
  - requirements.txt (dependencies)
  - static/ folder (CSS, JS, images)
  - templates/ folder (HTML templates)
  - mysite.py (WSGI entry point)
  - .env file (environment variables - keep this secure!)

## On PythonAnywhere Dashboard:

### 1. Create New Web App:
- [ ] Go to "Web" tab
- [ ] Click "Add a new web app"
- [ ] Choose "Flask" as framework
- [ ] Select Python version (3.9 recommended): Python 3.9 (Flask 3.0.3)
- [ ] For quickstart, use path: `/home/your-username/mysite/flask_app.py`

### 2. Upload Files:
- [ ] Go to "Files" tab
- [ ] Upload all your project files to `/home/your-username/myerp-app/`:
  - app.py
  - requirements.txt
  - static/ folder
  - templates/ folder
  - mysite.py

### 3. Set Up Virtual Environment:
- [ ] Open Bash console from PythonAnywhere dashboard
- [ ] Create virtual environment:
  ```bash
  mkvirtualenv myerp --python=/usr/bin/python3.9
  ```
- [ ] Activate it:
  ```bash
  workon myerp
  ```
- [ ] Navigate to your project:
  ```bash
  cd ~/myerp-app
  ```
- [ ] Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

### 4. Configure WSGI File:
- [ ] Click on "WSGI configuration file" link in Web tab
- [ ] Replace contents with:
```python
import sys
import os

# Add your project directory to sys.path
path = '/home/your-username/myerp-app'
if path not in sys.path:
    sys.path.insert(0, path)

# Change to your project directory
os.chdir(path)

# Activate your virtual environment
activate_this = '/home/your-username/.virtualenvs/myerp/bin/activate_this.py'
exec(open(activate_this).read(), {'__file__': activate_this})

# Import your Flask application
from app import app as application

if __name__ == "__main__":
    application.run()
```

### 5. Set Environment Variables:
- [ ] In the WSGI file or in your app.py, set:
  - SECRET_KEY
  - GOOGLE_SHEET_ID
  - GOOGLE_SHEETS_CREDENTIALS_JSON

### 6. Reload Application:
- [ ] Click "Reload" button in Web tab

### 7. Test Your Deployment:
- [ ] Visit your site at: https://your-username.pythonanywhere.com
- [ ] Test admin login
- [ ] Test student access
- [ ] Verify Google Sheets integration works

## Troubleshooting:
- Check error logs in Web tab if issues occur
- Ensure all file paths are correct
- Verify dependencies are installed in virtual environment
- Confirm environment variables are set correctly

## Once Successful:
- [ ] Share the URL: https://your-username.pythonanywhere.com with users
- [ ] Students can now access from home!