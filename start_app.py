#!/usr/bin/env python3
"""
Start script for the ERP application with background sync worker
"""

import os
from app import app
from sync_utils import start_background_sync

if __name__ == "__main__":
    # Start background sync worker
    try:
        start_background_sync()
        print("Background sync worker started")
    except Exception as e:
        print(f"Could not start background sync worker: {e}")

    # Run the Flask app
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 8080)))