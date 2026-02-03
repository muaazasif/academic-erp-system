@echo off
REM Deployment script for Academic ERP System

echo Preparing Academic ERP System for GitHub deployment...

REM Create a clean directory for the repository
mkdir academic-erp-system
cd academic-erp-system

REM Copy all necessary files from the source directory
copy ..\app.py .
copy ..\requirements.txt .
xcopy ..\static static /E /I
xcopy ..\templates templates /E /I
copy ..\README.md .
copy ..\LICENSE .
copy ..\.gitignore .
copy ..\Dockerfile .
copy ..\railway.json .
copy ..\mysite.py .
copy ..\CLOUD_RUN_DEPLOYMENT.md .
copy ..\PYTHONANYWHERE_DEPLOYMENT.md .
copy ..\RAILWAY_DEPLOYMENT.md .
copy ..\PYTHONANYWHERE_DEPLOYMENT_CHECKLIST.md .
copy ..\GLITCH_DEPLOYMENT.md .

echo Files copied successfully!

echo Initializing Git repository...
git init

echo Adding files to repository...
git add .

echo Creating initial commit...
git commit -m "Initial commit: Academic ERP System with Student Management, Attendance Tracking, Quizzes, Assignments, and Mid-term Examinations"

echo Repository prepared successfully!
echo.
echo To complete the GitHub deployment:
echo 1. Create a repository on GitHub: https://github.com/muaazasif/academic-erp-system
echo 2. Run these commands in the academic-erp-system directory:
echo    git remote add origin https://github.com/muaazasif/academic-erp-system.git
echo    git branch -M main
echo    git push -u origin main
echo.
echo After pushing to GitHub, you can deploy to various platforms using the guides provided.
pause