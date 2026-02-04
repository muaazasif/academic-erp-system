#!/usr/bin/env python3
"""
Migration script to recreate the database with location fields in the Attendance table
"""

import os
import sys

def migrate_database():
    """Recreate the database with proper schema including location fields"""
    try:
        # Import the app and database
        from app import app, db, Admin
        
        # Create all tables (this will include the location fields that are already defined in the model)
        with app.app_context():
            # Drop all tables first to ensure clean state
            print("Dropping existing tables...")
            db.drop_all()
            print("Creating tables with updated schema...")
            # Create all tables with the updated schema
            db.create_all()
            
            # Recreate the default admin user
            admin = Admin.query.filter_by(username='admin').first()
            if not admin:
                admin = Admin(username='admin')
                admin.set_password('admin123')
                db.session.add(admin)
                db.session.commit()
                print("Created default admin user")
        
        print("Database recreated with location-based attendance fields!")

        # Verify the attendance table has the location columns
        with app.app_context():  # Ensure we're in the app context
            from sqlalchemy import inspect
            inspector = inspect(db.engine)
            columns = [col['name'] for col in inspector.get_columns('attendance')]
            print(f"Columns in attendance table: {columns}")

            expected_cols = ['check_in_location', 'check_out_location']
            missing_cols = [col for col in expected_cols if col not in columns]

            if missing_cols:
                print(f"ERROR: Missing columns: {missing_cols}")
                return False
            else:
                print("SUCCESS: All location columns are present in the attendance table")
                return True
            
    except Exception as e:
        print(f"Error during migration: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Starting database migration for location-based attendance...")
    success = migrate_database()
    if success:
        print("\n\xE2\x9C\x93 Database migration completed successfully!")
    else:
        print("\n\xE2\x9C\x97 Database migration failed!")
        sys.exit(1)