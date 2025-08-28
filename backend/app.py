"""
Main FastAPI application factory and configuration.
"""
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import logging
import sys
import os
import socket
from dotenv import load_dotenv

def detect_backend_environment() -> str:
    """
    Detect the current backend environment based on various indicators.
    Returns 'development', 'staging', or 'production'.
    """
    # Check environment variable first (most explicit)
    env = os.getenv('ENVIRONMENT', '').lower()
    if env in ['development', 'dev', 'local']:
        return 'development'
    elif env in ['staging', 'stage', 'test']:
        return 'staging'
    elif env in ['production', 'prod', 'live']:
        return 'production'
    
    # Check hostname patterns as fallback
    try:
        hostname = socket.gethostname().lower()
        if 'vs81t' in hostname or 'staging' in hostname:
            return 'staging'
        elif 'vs84' in hostname or 'prod' in hostname:
            return 'production'
    except:
        pass
    
    # Default to development for safety (uses mock data)
    return 'development'

# Detect environment and load appropriate .env file
environment = detect_backend_environment()
env_file = f'.env.{environment}'
env_path = os.path.join(os.path.dirname(__file__), env_file)

if os.path.exists(env_path):
    load_dotenv(env_path)
    print(f"Loaded environment configuration from {env_file}")
else:
    print(f"Warning: Environment file {env_file} not found, using system environment variables")
    load_dotenv()  # Fallback to default .env file

# Add the src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.infrastructure.config.app_config import AppConfig
from src.presentation.controllers.raw_data_controller import raw_data_router
from src.presentation.controllers.market_data_controller import market_data_router
from src.presentation.controllers.data_upload_controller import data_upload_router


def create_app() -> FastAPI:
    """Create and configure the FastAPI application."""
    app = FastAPI(
        title="Excel Backend API",
        version="1.0.0",
        description="Backend API for Prime Excellence Excel Add-in"
    )
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO if not AppConfig.DEBUG else logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Configure CORS
    app.add_middleware(
        CORSMiddleware,
        allow_origins=AppConfig.CORS_ORIGINS,
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    
    # Include routers
    app.include_router(raw_data_router)
    app.include_router(market_data_router)
    app.include_router(data_upload_router)
    
    # Health check endpoint
    @app.get('/api/health')
    async def health_check():
        """Health check endpoint."""
        return {
            'status': 'healthy',
            'message': 'Excel Backend API is running',
            'environment': AppConfig.ENVIRONMENT,
            'cors_origins': AppConfig.CORS_ORIGINS,
            'nifi_endpoint': AppConfig.NIFI_ENDPOINT,
            'debug_mode': AppConfig.DEBUG
        }
    
    # Debug endpoint for connectivity testing
    @app.get('/api/debug')
    async def debug_info():
        """Debug endpoint to check backend configuration."""
        return {
            'status': 'debug',
            'environment': AppConfig.ENVIRONMENT,
            'host': AppConfig.HOST,
            'port': AppConfig.PORT,
            'cors_origins': AppConfig.CORS_ORIGINS,
            'nifi_endpoint': AppConfig.NIFI_ENDPOINT,
            'nifi_verify_ssl': AppConfig.NIFI_VERIFY_SSL,
            'debug_mode': AppConfig.DEBUG,
            'endpoints': [
                '/api/health',
                '/api/debug',
                '/api/raw-data/categories',
                '/api/raw-data/funds/{catalog}',
                '/api/raw-data/download',
                '/api/market-data/securities',
                '/api/market-data/fields/{security}',
                '/api/market-data/download',
                '/api/data-upload/upload',
                '/api/data-upload/types',
                '/api/data-upload/status/{upload_id}'
            ]
        }
    
    # Root endpoint
    @app.get('/')
    async def root():
        """Root endpoint."""
        return {
            'message': 'Excel Backend API',
            'version': '1.0.0',
            'endpoints': [
                '/api/health',
                '/api/raw-data/categories',
                '/api/raw-data/funds/{catalog}',
                '/api/raw-data/download',
                '/api/market-data/securities',
                '/api/market-data/fields/{security}',
                '/api/market-data/download',
                '/api/data-upload/upload',
                '/api/data-upload/types',
                '/api/data-upload/status/{upload_id}'
            ]
        }
    
    # Global exception handler for 404
    from fastapi import Request
    from fastapi.responses import JSONResponse
    
    @app.exception_handler(404)
    async def not_found_handler(request: Request, exc):
        """Handle 404 errors."""
        return JSONResponse(
            status_code=404,
            content={
                'success': False,
                'error': 'Endpoint not found'
            }
        )
    
    # Global exception handler for 500
    @app.exception_handler(500)
    async def internal_error_handler(request: Request, exc):
        """Handle 500 errors."""
        return JSONResponse(
            status_code=500,
            content={
                'success': False,
                'error': 'Internal server error'
            }
        )
    
    return app

def main():
    """Main entry point for the application."""
    import uvicorn
    
    logger = logging.getLogger(__name__)
    logger.info(f"Starting Excel Backend API on {AppConfig.HOST}:{AppConfig.PORT}")
    logger.info(f"Debug mode: {AppConfig.DEBUG}")
    logger.info(f"CORS origins: {AppConfig.CORS_ORIGINS}")
    
    # Log SSL configuration for NiFi
    logger.info(f"NiFi endpoint: {AppConfig.NIFI_ENDPOINT}")
    logger.info(f"NiFi SSL verification: {AppConfig.NIFI_VERIFY_SSL}")
    
    ssl_config = AppConfig.get_nifi_ssl_config()
    if AppConfig.NIFI_VERIFY_SSL:
        verify_setting = ssl_config.get('verify', 'system default')
        cert_setting = ssl_config.get('cert', 'not configured')
        logger.info(f"NiFi SSL verify setting: {verify_setting}")
        logger.info(f"NiFi client certificate: {cert_setting}")
        
        # Check certificate file existence
        certificates_path = AppConfig.get_certificates_path()
        logger.info(f"Certificates directory: {certificates_path}")
        
        if AppConfig.NIFI_CA_CERT_PATH:
            ca_cert_path = os.path.join(certificates_path, AppConfig.NIFI_CA_CERT_PATH)
            if os.path.exists(ca_cert_path):
                logger.info(f"CA certificate found: {ca_cert_path}")
            else:
                logger.warning(f"CA certificate not found: {ca_cert_path}")
        
        if AppConfig.NIFI_CLIENT_CERT_PATH:
            client_cert_path = os.path.join(certificates_path, AppConfig.NIFI_CLIENT_CERT_PATH)
            if os.path.exists(client_cert_path):
                logger.info(f"Client certificate found: {client_cert_path}")
            else:
                logger.warning(f"Client certificate not found: {client_cert_path}")
    else:
        logger.warning("SSL verification is DISABLED for NiFi connections - not recommended for production")
    
    # Create app
    app = create_app()
    
    # Run with uvicorn
    uvicorn.run(
        app,
        host=AppConfig.HOST,
        port=AppConfig.PORT,
        log_level="debug" if AppConfig.DEBUG else "info"
    )


if __name__ == '__main__':
    main()