# Excel Add-in Backend Windows Service Wrapper Script
# This script wraps the Flask backend for Windows Service execution

import os
import sys
import logging
import traceback
from pathlib import Path

# Set up logging for Windows Service
def setup_logging():
    """Configure logging for Windows Service"""
    log_dir = Path("C:/Logs/ExcelAddin")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Configure logging with both file and console output
    logging.basicConfig(
        level=logging.DEBUG,  # More verbose logging for debugging
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / 'backend-service.log', mode='a', encoding='utf-8'),
            logging.StreamHandler(sys.stdout),
            logging.StreamHandler(sys.stderr)  # Also log to stderr for NSSM capture
        ]
    )
    
    # Also set up werkzeug (Flask) logging
    werkzeug_logger = logging.getLogger('werkzeug')
    werkzeug_logger.setLevel(logging.INFO)
    
    logger = logging.getLogger(__name__)
    logger.info("=" * 50)
    logger.info("Starting Excel Add-in Backend Service")
    logger.info("Python version: %s", sys.version)
    logger.info("Python executable: %s", sys.executable)
    logger.info("Current working directory: %s", os.getcwd())
    logger.info("Script path: %s", __file__)
    logger.info("Python path: %s", sys.path)
    logger.info("=" * 50)
    return logger

def main():
    """Main entry point for Windows Service"""
    logger = None
    try:
        # Set up logging first
        logger = setup_logging()
        
        # Get the backend directory from script location
        script_path = Path(__file__).resolve()
        backend_dir = script_path.parent
        logger.info("Script location: %s", script_path)
        logger.info("Backend directory: %s", backend_dir)
        
        # Change to backend directory
        original_cwd = os.getcwd()
        os.chdir(backend_dir)
        logger.info("Changed working directory from %s to %s", original_cwd, os.getcwd())
        
        # Add backend directory to Python path (at the beginning)
        backend_str = str(backend_dir)
        if backend_str not in sys.path:
            sys.path.insert(0, backend_str)
            logger.info("Added %s to Python path", backend_str)
        
        # Set production environment variables
        env_vars = {
            'FLASK_ENV': 'production',
            'DEBUG': 'false',
            'HOST': '127.0.0.1',
            'PORT': '5000'
        }
        
        for key, value in env_vars.items():
            os.environ.setdefault(key, value)
            logger.info("Environment variable %s = %s", key, os.environ.get(key))
        
        # Log all environment variables for debugging
        logger.info("All environment variables:")
        for key in sorted(os.environ.keys()):
            if any(pattern in key.upper() for pattern in ['PYTHON', 'FLASK', 'DEBUG', 'HOST', 'PORT', 'PATH']):
                logger.info("  %s = %s", key, os.environ[key])
        
        # Check if required files exist
        required_files = ['app.py', 'requirements.txt']
        for file_name in required_files:
            file_path = backend_dir / file_name
            if file_path.exists():
                logger.info("✓ Found required file: %s", file_path)
            else:
                logger.error("✗ Missing required file: %s", file_path)
        
        # Test imports before running Flask
        logger.info("Testing imports...")
        try:
            logger.info("Importing Flask...")
            import flask
            logger.info("✓ Flask imported successfully (version: %s)", flask.__version__)
        except ImportError as e:
            logger.error("✗ Failed to import Flask: %s", e)
            logger.error("Try running: pip install flask")
            raise
        
        try:
            logger.info("Importing app module...")
            from app import create_app
            logger.info("✓ app module imported successfully")
        except ImportError as e:
            logger.error("✗ Failed to import app module: %s", e)
            logger.error("Ensure app.py exists in %s", backend_dir)
            logger.error("Current directory contents: %s", list(os.listdir(backend_dir)))
            raise
        
        # Create Flask app
        logger.info("Creating Flask application...")
        app = create_app()
        logger.info("✓ Flask app created successfully")
        
        # Log Flask app configuration
        logger.info("Flask app configuration:")
        for key, value in app.config.items():
            if 'SECRET' not in key and 'PASSWORD' not in key:
                logger.info("  %s = %s", key, value)
        
        # Get host and port from environment
        host = os.environ.get('HOST', '127.0.0.1')
        port = int(os.environ.get('PORT', 5000))
        
        logger.info("Starting Flask development server...")
        logger.info("Host: %s", host)
        logger.info("Port: %s", port)
        logger.info("Debug: %s", False)
        logger.info("Threaded: %s", True)
        
        # Run the Flask development server
        # Note: For production, consider using Gunicorn or similar WSGI server
        app.run(
            host=host,
            port=port,
            debug=False,
            threaded=True,
            use_reloader=False  # Disable reloader for service compatibility
        )
        
    except KeyboardInterrupt:
        if logger:
            logger.info("Service stopped by KeyboardInterrupt")
    except Exception as e:
        if not logger:
            # Fallback logging if logger setup failed
            import logging
            logging.basicConfig(level=logging.ERROR)
            logger = logging.getLogger(__name__)
        
        logger.error("Service failed to start: %s", str(e))
        logger.error("Exception type: %s", type(e).__name__)
        logger.error("Full traceback:")
        logger.error(traceback.format_exc())
        
        # Also print to stderr for immediate visibility
        print(f"FATAL ERROR: Service failed to start: {e}", file=sys.stderr)
        print(f"Exception type: {type(e).__name__}", file=sys.stderr)
        print("Full traceback:", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        
        # Exit with error code
        sys.exit(1)

if __name__ == '__main__':
    main()