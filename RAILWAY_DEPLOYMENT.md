# Deploy to Railway

## Students ghar se bhi access kar sakte hain! 🏠

### Kaise?

App deploy on Railway

URL milega:

```
https://your-project-name.up.railway.app
```

## Benefits:

✔️ 100+ users easy  
✔️ Auto scale  
✔️ Free tier available  
✔️ Students ghar se access kar sakte hain  
✔️ No server management needed  
✔️ SSL certificate automatically  
✔️ Modern deployment platform  
✔️ Integrated with Google services  

## How to Deploy on Railway

### Step 1: Create Account
1. Go to [Railway.app](https://railway.app/)
2. Create an account using your email or GitHub
3. Log in to your dashboard

### Step 2: Create New Project
1. Click "New Project" 
2. Select "Empty Project" or "Deploy from Repository"
3. If choosing "Empty Project", you'll deploy by uploading files directly

### Step 3: Upload Your Files
You can deploy to Railway in several ways:

#### Option A: Deploy from Local Machine (Without GitHub)
1. Install Railway CLI:
   ```bash
   npm install -g @railway/cli
   ```
2. Login to Railway:
   ```bash
   railway login
   ```
3. Initialize your project:
   ```bash
   railway init
   ```
4. Link to your existing project or create new:
   ```bash
   railway link
   ```
5. Deploy your project:
   ```bash
   railway up
   ```

#### Option B: Upload via Dashboard
1. In your Railway dashboard, click "Deploy from repository"
2. Even without GitHub, you can upload a ZIP file of your project
3. Click "Deploy from ZIP" (if available) or use the CLI method above

#### Option C: Deploy via Railway Desktop App
1. Download Railway's desktop application
2. Connect your project folder
3. Click deploy

### Step 4: Configure Variables
After deployment, set these environment variables in Railway:

1. Go to your project in Railway dashboard
2. Click on "Variables" or "Settings"
3. Add these variables:
   - `SECRET_KEY`: your-very-secure-secret-key
   - `GOOGLE_SHEET_ID`: your-google-sheet-id
   - `GOOGLE_SHEETS_CREDENTIALS_JSON`: your-service-account-json

### Step 5: Configure Your Application
Railway will automatically detect and use:
- `Dockerfile` for containerization
- `requirements.txt` for Python dependencies
- `railway.json` for configuration

### Step 6: Wait for Deployment
1. Railway will build your container
2. This may take 2-5 minutes
3. Check the logs for any errors

### Step 7: Access Your Application
1. After successful deployment, you'll see your URL
2. It will look like: `https://your-project-name.up.railway.app`
3. Share this URL with students and admins

## Required Files for Railway

Your project must include:

### `Dockerfile` (already created)
Contains instructions for building your application container.

### `requirements.txt`
Contains all Python dependencies.

### `app.py`
Your main Flask application file.

### `railway.json`
Configuration file for Railway deployment.

### `static/` and `templates/` folders
All your front-end assets and HTML templates.

## Environment Variables Setup

On Railway, set these environment variables:

```
SECRET_KEY=your-very-secure-secret-key
GOOGLE_SHEET_ID=your-google-sheet-id
GOOGLE_SHEETS_CREDENTIALS_JSON={"your":"credentials"}
```

## Port Configuration

Railway automatically sets the PORT environment variable. Your app.py already handles this correctly with:
```python
port = int(os.environ.get('PORT', 5000))
```

## Database Configuration

For production on Railway, consider using PostgreSQL:

```python
import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

# Use Railway's PostgreSQL if available, otherwise fallback to SQLite
DATABASE_URL = os.environ.get('DATABASE_URL')
if DATABASE_URL:
    # Remove the "postgres://" prefix and replace with "postgresql://"
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://")
    app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
else:
    app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///erp_system.db'
```

## Post Deployment Steps

1. After deployment, Railway will show your app URL
2. It will look like: `https://your-project-name.up.railway.app`
3. Share this URL with students and admins
4. The app will be accessible 24/7

## Troubleshooting

### Common Issues:

1. **Build Failures**: Check that all dependencies are in requirements.txt
2. **Port Binding**: Make sure your app uses `process.env.PORT` or defaults to 5000
3. **Environment Variables**: Ensure all required variables are set
4. **File Structure**: Verify all files are in the correct location

### Checking Logs:
- In Railway dashboard, go to your project
- Click "Logs" to see real-time logs
- Look for any error messages during startup

## Security Notes

1. **Never commit credentials to public repositories**
2. **Use Railway's environment variables for sensitive data**
3. **Regularly update dependencies**

## Scaling Configuration

Railway auto-scales based on demand. You can adjust settings in your project dashboard:
- Memory allocation
- CPU allocation
- Number of instances

## Success!

After deployment, your students can access the ERP system from anywhere:
- Home
- College
- Mobile devices
- Anywhere with internet connection

The application will be available at your Railway URL and students can access it from home!