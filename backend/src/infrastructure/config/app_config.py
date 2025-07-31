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
    
    # SSL/HTTPS configuration for external services
    @classmethod
    def get_certificates_path(cls) -> str:
        """Get the path to the certificates directory."""
        # Get the backend directory path
        backend_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
        return os.path.join(backend_dir, 'certificates')
    
    # NiFi HTTPS Configuration
    NIFI_ENDPOINT = os.getenv('NIFI_ENDPOINT', 'https://server-vs81t.intranet.local:8443/nifi/api/excel-addin-upload')
    NIFI_VERIFY_SSL = os.getenv('NIFI_VERIFY_SSL', 'true').lower() == 'true'
    
    # Certificate file paths (relative to certificates directory)
    NIFI_CA_CERT_PATH = os.getenv('NIFI_CA_CERT_PATH', None)
    NIFI_CLIENT_CERT_PATH = os.getenv('NIFI_CLIENT_CERT_PATH', None)
    NIFI_CLIENT_KEY_PATH = os.getenv('NIFI_CLIENT_KEY_PATH', None)
    
    @classmethod
    def get_nifi_ssl_config(cls) -> dict:
        """
        Get SSL configuration for NiFi requests.
        Returns a dict with SSL settings for the requests library.
        """
        if not cls.NIFI_VERIFY_SSL:
            return {'verify': False}
        
        ssl_config = {}
        certificates_path = cls.get_certificates_path()
        
        # CA certificate for server verification
        if cls.NIFI_CA_CERT_PATH:
            ca_cert_path = os.path.join(certificates_path, cls.NIFI_CA_CERT_PATH)
            if os.path.exists(ca_cert_path):
                ssl_config['verify'] = ca_cert_path
            else:
                # Default to system CA bundle if custom CA cert not found
                ssl_config['verify'] = True
        else:
            # Try default CA cert path
            default_ca_path = os.path.join(certificates_path, 'nifi-ca-cert.pem')
            if os.path.exists(default_ca_path):
                ssl_config['verify'] = default_ca_path
            else:
                ssl_config['verify'] = True
        
        # Client certificate for mutual TLS (if configured)
        if cls.NIFI_CLIENT_CERT_PATH and cls.NIFI_CLIENT_KEY_PATH:
            client_cert_path = os.path.join(certificates_path, cls.NIFI_CLIENT_CERT_PATH)
            client_key_path = os.path.join(certificates_path, cls.NIFI_CLIENT_KEY_PATH)
            
            if os.path.exists(client_cert_path) and os.path.exists(client_key_path):
                ssl_config['cert'] = (client_cert_path, client_key_path)
        elif cls.NIFI_CLIENT_CERT_PATH:
            # Single file containing both cert and key
            client_cert_path = os.path.join(certificates_path, cls.NIFI_CLIENT_CERT_PATH)
            if os.path.exists(client_cert_path):
                ssl_config['cert'] = client_cert_path
        
        return ssl_config