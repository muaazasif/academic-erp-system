from app import app, db

# Create all tables within the application context
with app.app_context():
    # This will create the new CourseOutline table if it doesn't exist
    # Note: This will not delete existing data
    db.create_all()
    print("Database tables updated successfully! New CourseOutline table has been created.")