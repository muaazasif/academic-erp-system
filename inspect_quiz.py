import os
import sys
from datetime import datetime

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the app module to access the database models
from app import app, Quiz, QuizQuestion, QuizOption, QuizSubmission, QuizAnswer

def inspect_quiz_data():
    """Inspect the quiz data in the database"""
    with app.app_context():
        # Get the quiz
        quiz = Quiz.query.first()
        if not quiz:
            print("No quizzes found in the database")
            return
        
        print(f"Quiz: {quiz.title} (ID: {quiz.id})")
        print(f"Description: {quiz.description}")
        print(f"Created by: {quiz.created_by}")
        print()
        
        # Get questions for this quiz
        questions = QuizQuestion.query.filter_by(quiz_id=quiz.id).all()
        for question in questions:
            print(f"Question {question.question_number}: {question.question_text}")
            
            # Get options for this question
            options = QuizOption.query.filter_by(question_id=question.id).all()
            for i, option in enumerate(options, 1):
                status = "CORRECT" if option.is_correct else "incorrect"
                print(f"  Option {i}: {option.option_text} [{status}]")
            print()

if __name__ == "__main__":
    inspect_quiz_data()