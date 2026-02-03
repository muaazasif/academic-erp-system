from app import app, db, Admin, Student
from werkzeug.security import generate_password_hash

def create_initial_data():
    with app.app_context():
        # Check if admin user already exists
        admin_exists = Admin.query.filter_by(username='admin').first()
        
        if not admin_exists:
            # Create default admin user
            admin = Admin(
                username='admin',
                password_hash=generate_password_hash('admin123')  # Using the default password from README
            )
            db.session.add(admin)
            db.session.commit()
            print("Admin user created successfully!")
            print("Username: admin")
            print("Password: admin123")
        else:
            print("Admin user already exists.")
            
        # Optionally create a sample student
        student_exists = Student.query.filter_by(student_id='101').first()
        
        if not student_exists:
            student = Student(
                student_id='101',
                name='John Doe',
                password_hash=generate_password_hash('student123')
            )
            db.session.add(student)
            db.session.commit()
            print("Sample student created successfully!")
            print("Student ID: 101")
            print("Password: student123")
        else:
            print("Sample student already exists.")

if __name__ == '__main__':
    create_initial_data()