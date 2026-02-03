import sys
import os

# Add your project directory to sys.path
path = '/home/' + os.environ['USER'] + '/myerp-app'
if path not in sys.path:
    sys.path.insert(0, path)

# Change to your project directory
os.chdir(path)

# Import your Flask application
from app import app as application

# For Google Sheets integration
if __name__ == "__main__":
    application.run()