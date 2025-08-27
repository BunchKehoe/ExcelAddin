"""
ASGI Entry Point for deployment servers
This module provides the ASGI application entry point for hosting.
Works with uvicorn, gunicorn, and other ASGI servers.
"""
import os
import sys
from pathlib import Path

# Get the backend directory from script location
backend_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(backend_dir))

# Set environment variables for production
os.environ.setdefault('ENVIRONMENT', 'production')
os.environ.setdefault('DEBUG', 'false')

# Import and create the FastAPI application
from app import create_app

# Create the ASGI application for deployment
application = create_app()

if __name__ == '__main__':
    # For testing purposes only - Use uvicorn or other ASGI servers for deployment
    import uvicorn
    uvicorn.run(application, host='127.0.0.1', port=5000, log_level="info")