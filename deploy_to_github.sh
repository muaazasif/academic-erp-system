#!/bin/bash
# Deployment script for Academic ERP System

echo "Preparing Academic ERP System for GitHub deployment..."

# Create a clean directory for the repository
mkdir -p academic-erp-system
cd academic-erp-system

# Copy all necessary files from the source directory
cp -r ../app.py .
cp -r ../requirements.txt .
cp -r ../static ./static
cp -r ../templates ./templates
cp -r ../README.md .
cp -r ../LICENSE .
cp -r ../.gitignore .
cp -r ../Dockerfile .
cp -r ../railway.json .
cp -r ../mysite.py .
cp -r ../CLOUD_RUN_DEPLOYMENT.md .
cp -r ../PYTHONANYWHERE_DEPLOYMENT.md .
cp -r ../RAILWAY_DEPLOYMENT.md .
cp -r ../PYTHONANYWHERE_DEPLOYMENT_CHECKLIST.md .
cp -r ../GLITCH_DEPLOYMENT.md .

echo "Files copied successfully!"

echo "Initializing Git repository..."
git init

echo "Adding files to repository..."
git add .

echo "Creating initial commit..."
git commit -m "Initial commit: Academic ERP System with Student Management, Attendance Tracking, Quizzes, Assignments, and Mid-term Examinations"

echo "Repository prepared successfully!"
echo ""
echo "To complete the GitHub deployment:"
echo "1. Create a repository on GitHub: https://github.com/muaazasif/academic-erp-system"
echo "2. Run these commands in the academic-erp-system directory:"
echo "   git remote add origin https://github.com/muaazasif/academic-erp-system.git"
echo "   git branch -M main"
echo "   git push -u origin main"
echo ""
echo "After pushing to GitHub, you can deploy to various platforms using the guides provided."