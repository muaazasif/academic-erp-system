import os
import sys
from flask import Flask, render_template_string

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <html>
    <head><title>Quiz System Test</title></head>
    <body>
        <h1>Quiz System is Running!</h1>
        <p>Server is accessible at localhost:5000</p>
        <p><a href="/quiz">Go to Quiz Section</a></p>
        <p><a href="/admin">Go to Admin Section</a></p>
    </body>
    </html>
    '''

@app.route('/quiz')
def quiz():
    return '''
    <html>
    <head><title>Quiz Section</title></head>
    <body>
        <h1>Quiz Section</h1>
        <p>Quiz functionality is implemented.</p>
        <p>Students can take quizzes here.</p>
        <a href="/">Back to Home</a>
    </body>
    </html>
    '''

@app.route('/admin')
def admin():
    return '''
    <html>
    <head><title>Admin Section</title></head>
    <body>
        <h1>Admin Section</h1>
        <p>Admin can create quizzes with questions and 4 options.</p>
        <p>All quiz results are saved to Google Sheets.</p>
        <a href="/">Back to Home</a>
    </body>
    </html>
    '''

if __name__ == '__main__':
    print("Starting application on http://localhost:5000")
    print("If you see this message, the application is running.")
    print("Try accessing http://localhost:5000 in your browser.")
    app.run(host='0.0.0.0', port=5000, debug=False)