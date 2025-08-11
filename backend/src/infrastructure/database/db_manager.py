"""
Database connection and session management.
"""
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker, Session
from contextlib import contextmanager
from typing import Generator
import logging

from ..config.app_config import DatabaseConfig

logger = logging.getLogger(__name__)


class DatabaseManager:
    """Database connection and session management."""
    
    def __init__(self):
        self._config = DatabaseConfig()
        self._engine = None
        self._session_factory = None
        self._is_mock_mode = False
        self._initialize_database()
    
    def _initialize_database(self):
        """Initialize database engine and session factory."""
        database_url = self._config.database_url
        
        if not database_url:
            # No database configured - this is expected for local development
            self._is_mock_mode = True
            logger.info(f"No database configured for {self._config.environment} environment - will use mock repository")
            return
        
        try:
            self._engine = create_engine(
                database_url,
                echo=False,  # Set to True for SQL logging
                pool_pre_ping=True,
                pool_recycle=3600
            )
            self._session_factory = sessionmaker(bind=self._engine)
            logger.info(f"Database connection initialized successfully for {self._config.environment} environment")
            
        except Exception as e:
            logger.error(f"Failed to initialize database connection for {self._config.environment} environment: {e}")
            logger.info("Application will fall back to mock repository")
            self._is_mock_mode = True
            raise
    
    @property
    def is_mock_mode(self) -> bool:
        """Check if database is in mock mode (no real database connection)."""
        return self._is_mock_mode
    
    @contextmanager
    def get_session(self) -> Generator[Session, None, None]:
        """Get database session context manager."""
        if self._is_mock_mode:
            raise RuntimeError("Database is in mock mode - no real database connection available")
            
        session = self._session_factory()
        try:
            yield session
            session.commit()
        except Exception as e:
            session.rollback()
            logger.error(f"Database session error: {e}")
            raise
        finally:
            session.close()
    
    def execute_query(self, query: str, params: dict = None) -> list:
        """Execute raw SQL query and return results."""
        if self._is_mock_mode:
            raise RuntimeError("Database is in mock mode - no real database connection available")
            
        with self.get_session() as session:
            try:
                result = session.execute(text(query), params or {})
                
                # Convert result to list of dictionaries
                columns = result.keys()
                rows = result.fetchall()
                
                return [dict(zip(columns, row)) for row in rows]
                
            except Exception as e:
                logger.error(f"Query execution failed: {e}")
                logger.error(f"Query: {query}")
                logger.error(f"Parameters: {params}")
                raise
    
    def execute_scalar_query(self, query: str, params: dict = None):
        """Execute query and return single value."""
        if self._is_mock_mode:
            raise RuntimeError("Database is in mock mode - no real database connection available")
            
        with self.get_session() as session:
            try:
                result = session.execute(text(query), params or {})
                return result.scalar()
            except Exception as e:
                logger.error(f"Scalar query execution failed: {e}")
                raise


# Global database manager instance
db_manager = DatabaseManager()