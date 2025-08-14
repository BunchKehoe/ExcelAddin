"""
Simple runner script for the Flask backend.
"""
import os
import sys

# Add the backend directory to Python path
backend_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, backend_dir)

# Check for required dependencies before importing app
def check_dependencies():
    """Check if required dependencies are installed."""
    missing_packages = []
    
    try:
        import flask
    except ImportError:
        missing_packages.append('flask')
    
    try:
        import dotenv
    except ImportError:
        missing_packages.append('python-dotenv')
    
    try:
        import flask_cors
    except ImportError:
        missing_packages.append('flask-cors')
    
    if missing_packages:
        print("\n" + "="*60)
        print("ERROR: Required Python packages are not installed!")
        print("="*60)
        print(f"\nMissing packages: {', '.join(missing_packages)}")
        print("\nTo fix this issue, run the following commands from the backend directory:")
        print("  poetry install")
        print("  poetry shell")
        print("\nThen try running the application again:")
        print("  python run.py")
        print("\n" + "="*60)
        sys.exit(1)

# Check dependencies before importing app
check_dependencies()

from app import main

if __name__ == '__main__':
    main()