# Deployment Guide for ERP Application

## Option 1: Deploy to PythonAnywhere (Recommended for Flask apps)

1. Sign up at https://www.pythonanywhere.com/
2. Upload your project files
3. Create a virtual environment and install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Configure your WSGI file to run the Flask app
5. Set environment variables for your Google Sheets API

## Option 2: Deploy to Heroku

1. Create a Procfile:
   ```
   web: gunicorn app:app
   ```
2. Create runtime.txt:
   ```
   python-3.9.0
   ```
3. Push to Heroku:
   ```bash
   heroku create
   heroku buildpacks:set heroku/python
   git push heroku master
   ```

## Option 3: Deploy to Vercel (with serverless functions)

1. Install Vercel CLI:
   ```
   npm i -g vercel
   ```
2. Initialize and deploy:
   ```
   vercel
   ```

## Option 4: Deploy to Render

1. Create a render.yaml file:
   ```yaml
   services:
   - type: web
     name: erp-system
     env: python
     buildCommand: pip install -r requirements.txt
     startCommand: python app.py
     envVars:
     - key: SECRET_KEY
       value: your-secret-key
     - key: DATABASE_URL
       value: your-database-url
   ```

## Important Notes:

1. Your application uses SQLite database which may not persist on some platforms
2. You'll need to configure Google Sheets API credentials on the hosting platform
3. Consider using PostgreSQL instead of SQLite for production deployments
4. Make sure to set environment variables for sensitive data

## For Cloudflare Pages (Static Content Only):

Cloudflare Pages is designed for static sites, not dynamic Flask applications. You would need to convert your Flask app to a static site or use it as a frontend with an API backend hosted elsewhere.

## Recommended Approach:

For your ERP system with database and backend logic, I recommend deploying to:
- PythonAnywhere
- Heroku
- Render
- AWS Elastic Beanstalk
- Google Cloud Platform
- DigitalOcean App Platform

These platforms support Python Flask applications with databases and Google Sheets integration.