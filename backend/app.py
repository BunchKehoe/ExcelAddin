"""
Main Flask application factory and configuration.
"""
from flask import Flask, jsonify
from flask_cors import CORS
import logging
import sys
import os

# Add the src directory to Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.infrastructure.config.app_config import AppConfig
from src.presentation.controllers.raw_data_controller import raw_data_bp
from src.presentation.controllers.market_data_controller import market_data_bp


def create_app() -> Flask:
    """Create and configure the Flask application."""
    app = Flask(__name__)
    
    # Configure logging
    logging.basicConfig(
        level=logging.INFO if not AppConfig.DEBUG else logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Configure CORS
    CORS(app, origins=AppConfig.CORS_ORIGINS)
    
    # Register blueprints
    app.register_blueprint(raw_data_bp)
    app.register_blueprint(market_data_bp)
    
    # Health check endpoint
    @app.route('/api/health', methods=['GET'])
    def health_check():
        """Health check endpoint."""
        return jsonify({
            'status': 'healthy',
            'message': 'Excel Backend API is running'
        })
    
    # Root endpoint
    @app.route('/', methods=['GET'])
    def root():
        """Root endpoint."""
        return jsonify({
            'message': 'Excel Backend API',
            'version': '1.0.0',
            'endpoints': [
                '/api/health',
                '/api/raw-data/categories',
                '/api/raw-data/funds/<catalog>',
                '/api/raw-data/download',
                '/api/market-data/securities',
                '/api/market-data/fields/<security>',
                '/api/market-data/download'
            ]
        })
    
    # Global error handler
    @app.errorhandler(404)
    def not_found(error):
        """Handle 404 errors."""
        return jsonify({
            'success': False,
            'error': 'Endpoint not found'
        }), 404
    
    @app.errorhandler(500)
    def internal_error(error):
        """Handle 500 errors."""
        return jsonify({
            'success': False,
            'error': 'Internal server error'
        }), 500
    
    return app


def main():
    """Main entry point for the application."""
    app = create_app()
    
    logger = logging.getLogger(__name__)
    logger.info(f"Starting Excel Backend API on {AppConfig.HOST}:{AppConfig.PORT}")
    logger.info(f"Debug mode: {AppConfig.DEBUG}")
    logger.info(f"CORS origins: {AppConfig.CORS_ORIGINS}")
    
    app.run(
        host=AppConfig.HOST,
        port=AppConfig.PORT,
        debug=AppConfig.DEBUG
    )


if __name__ == '__main__':
    main()