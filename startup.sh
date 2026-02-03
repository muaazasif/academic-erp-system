#!/bin/bash

# Initialize the database if it doesn't exist
if [ ! -f "instance/erp_system.db" ]; then
    echo "Initializing database..."
    python init_db.py
    
    # Create initial admin user
    python create_initial_data.py
else
    echo "Database already exists."
fi

# Start the application
exec "$@"