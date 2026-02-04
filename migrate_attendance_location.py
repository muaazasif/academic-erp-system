#!/usr/bin/env python3
"""
Migration script to add location fields to the Attendance table
"""

import sqlite3
import sys
import os

def migrate_database():
    """Add location columns to the attendance table"""
    db_path = 'erp_system.db'

    # Connect to the database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    try:
        # Check if the attendance table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='attendance';")
        table_exists = cursor.fetchone() is not None

        if not table_exists:
            print("Attendance table does not exist. Creating the database schema first...")
            # Close current connection and use SQLAlchemy to create all tables
            conn.close()

            # Import and use SQLAlchemy to create tables
            from app import app, db
            with app.app_context():
                db.create_all()
                print("Database tables created successfully using SQLAlchemy")

            # Reconnect to the database
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

        # Check if the columns already exist
        cursor.execute("PRAGMA table_info(attendance)")
        columns = [column[1] for column in cursor.fetchall()]

        print(f"Current columns in attendance table: {columns}")

        # Add check_in_location column if it doesn't exist
        if 'check_in_location' not in columns:
            cursor.execute("ALTER TABLE attendance ADD COLUMN check_in_location TEXT")
            print("Added check_in_location column")
        else:
            print("check_in_location column already exists")

        # Add check_out_location column if it doesn't exist
        if 'check_out_location' not in columns:
            cursor.execute("ALTER TABLE attendance ADD COLUMN check_out_location TEXT")
            print("Added check_out_location column")
        else:
            print("check_out_location column already exists")

        # Commit the changes
        conn.commit()
        print("Database migration completed successfully!")

        # Verify the columns were added
        cursor.execute("PRAGMA table_info(attendance)")
        updated_columns = [column[1] for column in cursor.fetchall()]
        print(f"Updated columns in attendance table: {updated_columns}")

        return True
        
    except sqlite3.Error as e:
        print(f"SQLite error: {e}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False
    finally:
        conn.close()

if __name__ == "__main__":
    print("Starting database migration for location-based attendance...")
    success = migrate_database()
    if success:
        print("\n✓ Database migration completed successfully!")
    else:
        print("\n✗ Database migration failed!")
        sys.exit(1)