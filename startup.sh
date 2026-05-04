#!/bin/bash

# Ensure instance directory exists
mkdir -p instance

# Initialize the database
# If DATABASE_URL is set (PostgreSQL), we always run init scripts as they check for existence
# If using SQLite, we only run if the file doesn't exist
if [ -n "$DATABASE_URL" ]; then
    echo "🚀 Using PostgreSQL database. Ensuring tables and initial data exist..."
    python init_db.py
    python create_initial_data.py
elif [ ! -f "instance/erp_system.db" ]; then
    echo "🚀 Initializing SQLite database..."
    python init_db.py
    python create_initial_data.py
else
    echo "✅ SQLite database already exists."
fi

# Start the application with the arguments passed to this script
echo "🎬 Starting application with: $@"
exec "$@"