# Deploy to PythonAnywhere

## Students ghar se bhi access kar sakte hain! 🏠

### Kaise?

App deploy on PythonAnywhere

URL milega:

```
https://your-username.pythonanywhere.com
```

## Benefits:

✔️ 100+ users easy  
✔️ Auto scale (based on plan)  
✔️ Free tier available  
✔️ Students ghar se access kar sakte hain  
✔️ No server management needed  
✔️ SSL certificate available  
✔️ Python-optimized environment  
✔️ Integrated with Google services  

## How to Deploy on PythonAnywhere

### Step 1: Create Account
1. Go to [PythonAnywhere.com](https://www.pythonanywhere.com/)
2. Create a free account (or paid for production use)
3. Log in to your dashboard

### Step 2: Create a New Web App
1. Click on "Web" tab in your dashboard
2. Click "Add a new web app"
3. Choose "Flask" as your web framework
4. Select the Python version:
   - Python 3.9 (Flask 3.0.3) - Recommended for compatibility
   - Python 3.10 (Flask 3.0.3)
   - Python 3.11 (Flask 3.0.3)
   - Python 3.12 (Flask 3.0.3)
   - Python 3.13 (Flask 3.0.3)
5. For "Quickstart new Flask project":
   - Enter a path for a Python file to hold your Flask app
   - Example: `/home/your-username/mysite/flask_app.py`
   - If this file already exists, its contents will be overwritten with the new app
6. Use the default settings unless you have specific requirements

### Step 3: Configure Your Application
1. After creation, you'll see configuration options
2. In the "Code" section, update the Source Code directory to point to your project
3. In the "Virtualenv" section, set the path to your virtual environment

### Step 4: Upload Your Files
You have several options to upload your files:

#### Option A: Upload via Dashboard
1. Go to "Files" tab in PythonAnywhere
2. Upload all your project files:
   - `app.py` - Your main Flask application
   - `requirements.txt` - Python dependencies
   - `static/` folder - CSS, JavaScript, images
   - `templates/` folder - HTML templates
   - Any other necessary files

#### Option B: Clone from GitHub
1. In the "Files" tab, open a bash console
2. Clone your repository:
   ```bash
   git clone https://github.com/your-username/your-repo.git
   ```

#### Option C: Upload via FTP/SFTP
1. Use your preferred FTP client
2. Connect to ftp.pythonanywhere.com
3. Upload files to your home directory

### Step 5: Install Dependencies
1. Open a Bash console in PythonAnywhere
2. Navigate to your project directory
3. Create and activate a virtual environment:
   ```bash
   mkvirtualenv myerp --python=/usr/bin/python3.9
   workon myerp
   ```
4. Install your dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Step 6: Configure WSGI File
1. Go to your web app configuration page
2. Click on the "WSGI configuration file" link
3. Update the WSGI file to point to your Flask app:

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
```

### Step 7: Set Environment Variables
1. In your PythonAnywhere dashboard, go to "Account" → "API token"
2. Create an API token for secure access
3. For environment variables, you can either:
   - Add them to your WSGI file using `os.environ`
   - Or store sensitive data securely in a separate config file

Example of setting environment variables in WSGI:
```python
import os
os.environ['SECRET_KEY'] = 'your-secret-key'
os.environ['GOOGLE_SHEET_ID'] = 'your-google-sheet-id'
os.environ['GOOGLE_SHEETS_CREDENTIALS_JSON'] = 'your-json-string'
```

### Step 8: Reload Your Application
1. Go back to the "Web" tab
2. Click "Reload" button to restart your application
3. Check the error logs if anything goes wrong

### Step 9: Test Your Application
1. Visit your application at `https://your-username.pythonanywhere.com`
2. Test all functionality including:
   - Admin login
   - Student access
   - Google Sheets integration
   - All forms and submissions

## Required Files for PythonAnywhere

Your project must include:

### `requirements.txt`
Contains all Python dependencies.

### `app.py`
Your main Flask application file.

### `static/` and `templates/` folders
All your front-end assets and HTML templates.

## Environment Variables Setup

On PythonAnywhere, set these environment variables:

```
SECRET_KEY=your-very-secure-secret-key
GOOGLE_SHEET_ID=your-google-sheet-id
GOOGLE_SHEETS_CREDENTIALS_JSON={"your":"credentials"}
```

## Database Configuration

For production, consider using PostgreSQL instead of SQLite:

```python
# In your app.py
import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

# Use PostgreSQL in production, SQLite for development
if os.environ.get('PYTHONANYWHERE_SITE'):
    # Production settings
    app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://username:password@localhost/dbname'
else:
    # Development settings
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///erp_system.db'
```

## Post Deployment Steps

1. After deployment, PythonAnywhere will show your app URL
2. It will look like: `https://your-username.pythonanywhere.com`
3. Share this URL with students and admins
4. The app will be accessible 24/7

## Troubleshooting

### Common Issues:

1. **Import Errors**: Make sure all dependencies are installed in the virtual environment
2. **Path Issues**: Ensure your WSGI file has the correct path to your application
3. **Permissions**: Make sure files have correct read permissions
4. **Database**: SQLite files need write permissions in the correct location

### Checking Logs:
- In PythonAnywhere dashboard, go to "Web" tab
- Check "Error log" and "Access log" for debugging information
- Look for any error messages during startup

## Security Notes

1. **Never commit credentials to public repositories**
2. **Use environment variables for sensitive data**
3. **Regularly update dependencies**
4. **Consider using a production database (PostgreSQL)**

## Scaling Options

- Free account: Limited resources, good for testing
- Hobby account: More resources, suitable for small deployments
- Paid accounts: Higher limits, better performance, SSL certificates

## Success!

After deployment, your students can access the ERP system from anywhere:
- Home
- College
- Mobile devices
- Anywhere with internet connection

The application will be available at your PythonAnywhere URL and students can access it from home!