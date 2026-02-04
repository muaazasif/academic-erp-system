"""
Database migration script to add location fields to Attendance table
"""
import sqlite3
import sys
import os

# Add the application directory to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def migrate_database():
    print("Starting database migration to add location fields to Attendance table...")

    # Connect to the SQLite database in the instance folder
    conn = sqlite3.connect('instance/erp_system.db')
    cursor = conn.cursor()
    
    try:
        # Check if the columns already exist
        cursor.execute("PRAGMA table_info(attendance)")
        columns = [column[1] for column in cursor.fetchall()]

        print(f"Current columns in attendance table: {columns}")

        # Add check_in_location column if it doesn't exist
        if 'check_in_location' not in columns:
            cursor.execute("ALTER TABLE attendance ADD COLUMN check_in_location TEXT")
            print("Added check_in_location column to attendance table")
        else:
            print("check_in_location column already exists")

        # Add check_out_location column if it doesn't exist
        if 'check_out_location' not in columns:
            cursor.execute("ALTER TABLE attendance ADD COLUMN check_out_location TEXT")
            print("Added check_out_location column to attendance table")
        else:
            print("check_out_location column already exists")
        
        # Commit the changes
        conn.commit()
        print("Database migration completed successfully!")
        
    except sqlite3.Error as e:
        print(f"Database error: {e}")
        conn.rollback()
    except Exception as e:
        print(f"Error during migration: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    migrate_database()