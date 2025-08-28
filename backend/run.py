"""
Service runner script for the FastAPI backend.
This script is specifically designed to work with Windows services (NSSM).
"""
import os
import sys
import logging
from pathlib import Path

# Setup logging for service debugging
log_dir = Path("C:/Logs/ExcelAddin")
log_dir.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / "service-startup.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def setup_environment():
    """Setup environment for service execution."""
    logger.info("Setting up service environment...")
    
    # Get the backend directory from script location
    backend_dir = Path(__file__).resolve().parent
    logger.info(f"Backend directory: {backend_dir}")
    
    # Add backend directory to Python path
    sys.path.insert(0, str(backend_dir))
    logger.info(f"Added to Python path: {backend_dir}")
    
    # Change working directory to backend
    os.chdir(backend_dir)
    logger.info(f"Changed working directory to: {os.getcwd()}")
    
    # Log environment variables
    env_vars = ['ENVIRONMENT', 'PORT', 'HOST', 'DEBUG']
    for var in env_vars:
        value = os.getenv(var, 'NOT_SET')
        logger.info(f"Environment variable {var}: {value}")

def check_dependencies():
    """Check if required dependencies are installed."""
    logger.info("Checking dependencies...")
    missing_packages = []
    
    try:
        import fastapi
        logger.info(f"FastAPI version: {fastapi.__version__}")
    except ImportError as e:
        missing_packages.append('fastapi')
        logger.error(f"Failed to import fastapi: {e}")
    
    try:
        import uvicorn
        logger.info(f"Uvicorn version: {uvicorn.__version__}")
    except ImportError as e:
        missing_packages.append('uvicorn')
        logger.error(f"Failed to import uvicorn: {e}")
    
    try:
        import dotenv
        logger.info("python-dotenv imported successfully")
    except ImportError as e:
        missing_packages.append('python-dotenv')
        logger.error(f"Failed to import python-dotenv: {e}")
    
    if missing_packages:
        error_msg = f"Missing required packages: {', '.join(missing_packages)}"
        logger.error(error_msg)
        logger.error("To fix this issue, run the following commands from the backend directory:")
        logger.error("  poetry install")
        logger.error("  poetry shell")
        logger.error("Then try running the service again.")
        raise ImportError(error_msg)
    
    logger.info("All dependencies available")

def main():
    """Main entry point for the service."""
    try:
        logger.info("=== Starting FastAPI Backend Service ===")
        
        # Setup environment
        setup_environment()
        
        # Check dependencies
        check_dependencies()
        
        # Import and run the app
        logger.info("Importing app module...")
        from app import main as app_main
        
        logger.info("Starting FastAPI application...")
        app_main()
        
    except Exception as e:
        logger.error(f"Failed to start service: {e}", exc_info=True)
        sys.exit(1)

if __name__ == '__main__':
    main()