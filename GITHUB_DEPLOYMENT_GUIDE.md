# GitHub Deployment Guide for Academic ERP System

## Step-by-Step Instructions

### Step 1: Create GitHub Repository
1. Go to https://github.com
2. Click on the "+" icon in the top right corner
3. Select "New repository"
4. Fill in the details:
   - Repository name: academic-erp-system
   - Description: Academic ERP System with Student Management, Attendance Tracking, Quizzes, Assignments, and Mid-term Examinations
   - Public: ✓ (check this box)
   - Add a README file: ✗ (leave unchecked)
   - Add .gitignore: Select "Python"
   - License: Select "MIT License"
5. Click "Create repository"

### Step 2: Prepare Your Local Repository
Open Command Prompt or Terminal in your project directory:

```bash
# Navigate to your project directory
cd "E:\Governor Sindh Course\Application\myerp_app"

# Initialize git in your project
git init

# Add all files to the repository
git add .

# Create initial commit
git commit -m "Initial commit: Academic ERP System with Student Management, Attendance Tracking, Quizzes, Assignments, and Mid-term Examinations"
```

### Step 3: Connect to Your GitHub Repository
1. On your GitHub repository page, click the green "Code" button
2. Copy the HTTPS URL (it should look like: https://github.com/muaazasif/academic-erp-system.git)
3. In your Command Prompt/Terminal, run:

```bash
# Connect your local repository to GitHub
git remote add origin https://github.com/muaazasif/academic-erp-system.git

# Verify the connection
git remote -v
```

### Step 4: Push Your Code to GitHub
```bash
# Push your code to GitHub
git branch -M main
git push -u origin main
```

### Step 5: Verify Deployment
1. Refresh your GitHub repository page
2. You should see all your files uploaded
3. The README.md file will be displayed

## Troubleshooting

### If you get authentication errors:
1. You may need to create a Personal Access Token:
   - Go to GitHub Settings → Developer settings → Personal access tokens
   - Click "Generate new token"
   - Give it a name and select appropriate permissions
   - Copy the generated token
2. When prompted for password, use the token instead of your GitHub password

### If you get "non-fast-forward" errors:
```bash
# Pull any changes first
git pull origin main --allow-unrelated-histories
# Then push again
git push origin main
```

## After GitHub Deployment

Once your code is on GitHub, you can deploy to various platforms:

### Option 1: Railway
1. Go to https://railway.app
2. Click "New Project"
3. Select "Deploy from GitHub"
4. Connect to your academic-erp-system repository
5. Railway will automatically detect the Dockerfile and deploy

### Option 2: PythonAnywhere
1. Go to https://www.pythonanywhere.com
2. Create an account
3. Click "Web" tab → "Add a new web app"
4. Select "Flask" and your Python version
5. Choose "Clone from a Git repo"
6. Enter your GitHub repository URL

### Option 3: Google Cloud Run
1. Install Google Cloud SDK
2. Use the deployment instructions in CLOUD_RUN_DEPLOYMENT.md

## Verification Checklist
- [ ] GitHub repository created at https://github.com/muaazasif/academic-erp-system
- [ ] All project files uploaded to GitHub
- [ ] README.md displays properly
- [ ] requirements.txt is present
- [ ] All templates and static files are uploaded
- [ ] Deployment guides are available in the repository

## Important Notes
- Your project contains sensitive files like cookies.txt - make sure these are properly handled
- The .gitignore file should prevent sensitive data from being uploaded
- After deployment to any platform, remember to set environment variables for Google Sheets integration

## Need Help?
If you encounter any issues during the process, please share the error message and I can provide specific guidance.