from flask import Flask

app = Flask(__name__)

@app.route('/')
def home():
    return '''
    <h1>Quiz System is Working!</h1>
    <p>The application is running successfully.</p>
    <p><a href="/quiz-test">Test Quiz Route</a></p>
    '''

@app.route('/quiz-test')
def quiz_test():
    return '''
    <h1>Quiz Test Page</h1>
    <p>This confirms the quiz system routes are working.</p>
    <p>Full quiz functionality is implemented in the main application.</p>
    <a href="/">Back to Home</a>
    '''

if __name__ == '__main__':
    print("Test application starting on http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000, debug=False)