@echo off
REM Initialize the database if it doesn't exist
IF NOT EXIST "instance\erp_system.db" (
    echo Initializing database...
    python init_db.py
    
    REM Create initial admin user
    python create_initial_data.py
) ELSE (
    echo Database already exists.
)

REM Start the application
gunicorn --bind 0.0.0.0:8000 --workers 2 app:app