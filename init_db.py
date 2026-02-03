import os
from app import app, db

# Ensure the instance directory exists
os.makedirs(app.instance_path, exist_ok=True)

# Create all tables within the application context
with app.app_context():
    db.create_all()
    print("Database tables created successfully!")
    print(f"Database file location: {os.path.join(app.instance_path, 'erp_system.db')}")