"""
WSGI Entry Point for IIS
This module provides the WSGI application entry point for IIS hosting.
Replaces the NSSM service wrapper for direct IIS integration.
"""
import os
import sys
from pathlib import Path

# Get the backend directory from script location
backend_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(backend_dir))

# Set environment variables for production
os.environ.setdefault('FLASK_ENV', 'production')
os.environ.setdefault('DEBUG', 'false')

# Import and create the Flask application
from app import create_app

# Create the WSGI application for IIS
application = create_app()

if __name__ == '__main__':
    # For testing purposes only - IIS will use the application object above
    application.run(host='127.0.0.1', port=5000, debug=False)