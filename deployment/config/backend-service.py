#!/usr/bin/env python3
"""
NSSM service wrapper for ExcelAddin backend.
This script ensures proper service startup and environment configuration.
"""
import os
import sys
import time
import logging
from pathlib import Path

def setup_logging():
    """Configure logging for the service."""
    log_dir = Path("C:/Logs/ExcelAddin")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / "backend-service.log"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def setup_environment():
    """Setup environment for the backend service."""
    # Get the directory where this script is located
    script_dir = Path(__file__).parent
    backend_dir = script_dir.parent.parent / "backend"
    
    # Add backend directory to Python path
    sys.path.insert(0, str(backend_dir))
    
    # Set working directory to backend
    os.chdir(str(backend_dir))
    
    # Ensure .env file exists for staging environment
    env_file = backend_dir / ".env"
    staging_env = backend_dir / ".env.staging"
    
    if not env_file.exists() and staging_env.exists():
        import shutil
        shutil.copy(str(staging_env), str(env_file))
        logger.info(f"Copied {staging_env} to {env_file}")
    
    # Set environment variables
    os.environ["ENVIRONMENT"] = "staging"
    os.environ["PYTHONPATH"] = str(backend_dir)
    
    return backend_dir

def main():
    """Main service entry point."""
    logger = setup_logging()
    logger.info("ExcelAddin Backend Service Starting...")
    
    try:
        # Setup environment
        backend_dir = setup_environment()
        logger.info(f"Working directory: {backend_dir}")
        
        # Import and run the Flask app
        from app import create_app
        
        app = create_app()
        
        # Configure for production
        host = os.getenv("HOST", "127.0.0.1")
        port = int(os.getenv("PORT", "5000"))
        
        logger.info(f"Starting Flask app on {host}:{port}")
        
        # Run the application
        app.run(
            host=host,
            port=port,
            debug=False,
            threaded=True,
            use_reloader=False
        )
        
    except Exception as e:
        logger.error(f"Failed to start backend service: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()