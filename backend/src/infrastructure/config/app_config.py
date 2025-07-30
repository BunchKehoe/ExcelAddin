"""
Configuration management for the backend application.
"""
import configparser
import os
from typing import Optional


class DatabaseConfig:
    """Database configuration management."""
    
    def __init__(self, config_file: str = "database.cfg"):
        self.config_file = config_file
        self._config = None
        self._load_config()
    
    def _load_config(self):
        """Load configuration from file."""
        self._config = configparser.ConfigParser()
        # Go up to the backend directory from src/infrastructure/config/
        backend_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
        config_path = os.path.join(backend_dir, self.config_file)
        
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"Configuration file not found: {config_path}")
        
        self._config.read(config_path)
    
    @property
    def database_url(self) -> str:
        """Get database connection URL."""
        if not self._config.has_section('database'):
            raise ValueError("Database section not found in configuration")
        
        return self._config.get('database', 'url')


class AppConfig:
    """Application configuration."""
    
    DEBUG = os.getenv('DEBUG', 'False').lower() == 'true'
    HOST = os.getenv('HOST', '0.0.0.0')
    PORT = int(os.getenv('PORT', '5000'))
    
    # CORS settings
    CORS_ORIGINS = os.getenv('CORS_ORIGINS', 'http://localhost:3000').split(',')