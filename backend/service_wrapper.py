# Excel Add-in Backend Windows Service Wrapper Script
# This script wraps the Flask backend for Windows Service execution

import os
import sys
import logging
from pathlib import Path

# Set up logging for Windows Service
def setup_logging():
    """Configure logging for Windows Service"""
    log_dir = Path("C:/Logs/ExcelAddin")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / 'backend-service.log'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("Starting Excel Add-in Backend Service")
    return logger

def main():
    """Main entry point for Windows Service"""
    try:
        logger = setup_logging()
        
        # Get the backend directory
        backend_dir = Path(__file__).parent
        os.chdir(backend_dir)
        
        # Add backend directory to Python path
        sys.path.insert(0, str(backend_dir))
        
        # Set production environment variables
        os.environ.setdefault('FLASK_ENV', 'production')
        os.environ.setdefault('DEBUG', 'false')
        os.environ.setdefault('HOST', '127.0.0.1')
        os.environ.setdefault('PORT', '5000')
        
        logger.info(f"Working directory: {backend_dir}")
        logger.info(f"Environment: {os.environ.get('FLASK_ENV')}")
        logger.info(f"Debug: {os.environ.get('DEBUG')}")
        logger.info(f"Host: {os.environ.get('HOST')}")
        logger.info(f"Port: {os.environ.get('PORT')}")
        
        # Import and run the Flask app
        from app import create_app
        
        app = create_app()
        logger.info("Flask app created successfully")
        
        # Run the Flask development server
        # Note: For production, consider using Gunicorn or similar WSGI server
        app.run(
            host=os.environ.get('HOST', '127.0.0.1'),
            port=int(os.environ.get('PORT', 5000)),
            debug=False,
            threaded=True
        )
        
    except Exception as e:
        logger = logging.getLogger(__name__)
        logger.error(f"Service failed to start: {str(e)}", exc_info=True)
        raise

if __name__ == '__main__':
    main()