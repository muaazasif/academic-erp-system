# Deployment to Google Cloud Run (RECOMMENDED!)

## Students ghar se bhi access kar sakte hain! 🏠

### Kaise?

App deploy on Google Cloud Run

URL milega:

```
https://myerp-app-xyz.a.run.app
```

## Benefits:

✔️ 100+ users easy
✔️ Auto scale
✔️ Free tier mein mostly chal jata hai
✔️ Students ghar se access kar sakte hain
✔️ High performance
✔️ SSL certificate automatically
✔️ Global availability
✔️ Production-ready environment
✔️ Google's infrastructure  

## Step-by-Step Deployment Guide

### Prerequisites:
1. Google Cloud Account
2. Google Cloud SDK installed
3. Billing enabled on your Google Cloud account (required for Cloud Run)

### Step 1: Setup Google Cloud SDK
```bash
# Download and install Google Cloud SDK
# Windows: https://dl.google.com/dl/cloudsdk/channels/rapid/GoogleCloudSDKInstaller.exe
# Mac/Linux: curl https://sdk.cloud.google.com | bash
```

### Step 2: Login to Google Cloud
```bash
gcloud auth login
```

### Step 3: Create a Project
```bash
gcloud projects create your-project-id
gcloud config set project your-project-id
```

### Step 4: Enable Required APIs
```bash
gcloud services enable run.googleapis.com
gcloud services enable containerregistry.googleapis.com
gcloud services enable sheets.googleapis.com
gcloud services enable drive.googleapis.com
```

### Step 5: Prepare Your Application Files

Make sure you have these files in your project:
- `app.py` - Your Flask application
- `requirements.txt` - Dependencies
- `Dockerfile` - Container configuration
- `main.py` - Entry point
- `app.yaml` - App Engine configuration (optional for Cloud Run)

### Step 6: Create Google Service Account for Google Sheets API

1. Go to Google Cloud Console
2. Navigate to IAM & Admin > Service Accounts
3. Create a new service account
4. Download the JSON key file
5. Share your Google Sheet with the service account email

### Step 7: Build and Deploy

#### Option A: Using Cloud Build (Recommended)
```bash
# Submit build to Cloud Build
gcloud builds submit --config cloudbuild.yaml .
```

#### Option B: Manual Deployment
```bash
# Build the container image
gcloud builds submit --tag gcr.io/YOUR_PROJECT_ID/myerp-app

# Deploy to Cloud Run
gcloud run deploy myerp-app \
  --image gcr.io/YOUR_PROJECT_ID/myerp-app \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --set-env-vars SECRET_KEY=your-very-secure-secret-key,GOOGLE_SHEET_ID=your-google-sheet-id \
  --set-secrets=GOOGLE_SHEETS_CREDENTIALS_JSON=google-sheets-credentials:latest
```

### Step 8: Access Your Application

After successful deployment, you'll receive a URL like:
```
https://myerp-app-xyz-a.run.app
```

Share this URL with students and admins to access the application from anywhere!

## Environment Variables Setup

You need to set these environment variables:

### SECRET_KEY
- Generate a strong secret key for Flask sessions
- Example: `openssl rand -hex 32`

### GOOGLE_SHEET_ID
- Get this from your Google Sheet URL
- Format: `https://docs.google.com/spreadsheets/d/[THIS_IS_YOUR_SHEET_ID]/edit`

### GOOGLE_SHEETS_CREDENTIALS_JSON
- Store your service account credentials as a secret in Google Secret Manager
- Format: JSON string of your service account key

## Security Best Practices

1. **Never commit credentials to source code**
2. **Use Google Secret Manager for sensitive data**
3. **Enable authentication if needed for your use case**
4. **Regularly rotate your secret keys**

## Troubleshooting

### Common Issues:

1. **Permission Errors**: Ensure your service account has Editor role on the project
2. **Build Failures**: Check that all dependencies are in requirements.txt
3. **API Access**: Verify that required APIs are enabled
4. **Billing**: Ensure billing is enabled for your project

### Testing Locally Before Deployment:

```bash
# Install dependencies
pip install -r requirements.txt

# Run locally
python app.py
```

## Scaling Configuration

By default, Cloud Run auto-scales based on demand. You can customize scaling:

```bash
gcloud run deploy myerp-app \
  --image gcr.io/YOUR_PROJECT_ID/myerp-app \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --concurrency 80 \
  --memory 512Mi \
  --cpu 1
```

## Monitoring

Monitor your application using Google Cloud Console:
- Visit: https://console.cloud.google.com/run
- Check logs: https://console.cloud.google.com/logs
- Monitor usage: https://console.cloud.google.com/monitoring

## Cost Optimization

- Under free tier: 2 million requests/month, 360 CPU hours/month, 5GB storage
- Scale to zero when not in use (no cost when idle)
- Monitor usage to stay within free limits

## Success!

After deployment, your students can access the ERP system from anywhere:
- Home
- College
- Mobile devices
- Anywhere with internet connection

The application will automatically scale to handle multiple users simultaneously!