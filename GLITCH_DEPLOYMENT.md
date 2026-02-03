# Deploy to Replit

## Students ghar se bhi access kar sakte hain! 🏠

### Kaise?

App deploy on Replit

URL milega:

```
https://your-project-name.your-username.repl.co
```

## Benefits:

✔️ 100+ users easy
✔️ Auto scale
✔️ Free hosting (with paid upgrade options)
✔️ Students ghar se access kar sakte hain
✔️ No server management needed
✔️ Real-time collaboration
✔️ SSL certificate automatically

## How to Deploy on Replit

### Method 1: Import from GitHub
1. Go to [Replit.com](https://replit.com/)
2. Create an account and click "Create"
3. Select "Import from GitHub"
4. Enter your GitHub repository URL
5. Replit will automatically detect and run your app

### Method 2: Create New Python Repl
1. Go to [Replit.com](https://replit.com/)
2. Click "Create" → "Python" template
3. Replace all files with your application files
4. Make sure you have:
   - `app.py` (your Flask app)
   - `requirements.txt` (Python dependencies)
   - `static/` folder (CSS, JS, images)
   - `templates/` folder (HTML templates)

### Method 3: Direct Upload
1. Go to [Replit.com](https://replit.com/)
2. Click "Create" → "Python"
3. Upload or copy-paste your code files
4. Make sure all files are uploaded correctly

## Required Files for Replit

Your project must include:

### `requirements.txt`
Contains all Python dependencies.

### `main.py` or `app.py`
Your main Flask application file.

### Environment Variables Setup

On Replit, set these environment variables in Secrets/Environment Variables:

```
SECRET_KEY=your-very-secure-secret-key
GOOGLE_SHEET_ID=your-google-sheet-id
GOOGLE_SHEETS_CREDENTIALS_JSON={"your":"credentials"}
```

## Post Deployment Steps

1. After deployment, Replit will show your app URL
2. It will look like: `https://your-project-name.your-username.repl.co`
3. Share this URL with students and admins
4. The app will be accessible 24/7 (unless project goes to sleep)

## Keeping Your App Awake

Free Replit projects may sleep after inactivity. To keep it awake:
1. Use UptimeRobot (free) to ping your app every 5 minutes
2. Set up a simple monitor to hit your app's URL regularly

## Troubleshooting

### Common Issues:

1. **App not starting**: Check the Shell/Console tab in Replit for error messages
2. **Port binding**: Make sure your app uses `os.environ.get('PORT', 5000)` for port detection
3. **Dependencies**: Run `pip install -r requirements.txt` in the shell
4. **File permissions**: Make sure all files are readable

### Checking Logs:
- In Replit editor, check the Shell/Console tab to see real-time logs
- Look for any error messages during startup

## Security Notes

1. **Never commit credentials to public repositories**
2. **Use Replit's secret environment variables for sensitive data**
3. **Regularly update dependencies**

## Success!

After deployment, your students can access the ERP system from anywhere:
- Home
- College
- Mobile devices
- Anywhere with internet connection

The application will be available at your Replit URL and students can access it from home!