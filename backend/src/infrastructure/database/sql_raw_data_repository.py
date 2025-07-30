"""
SQL Server implementation of raw data repository.
"""
from typing import List
import logging

from ...domain.repositories.raw_data_repository import IRawDataRepository
from ...domain.entities.raw_data import FileCategory, Fund, RawDataRecord, RawDataRequest
from ..database.db_manager import db_manager

logger = logging.getLogger(__name__)


class SqlRawDataRepository(IRawDataRepository):
    """SQL Server implementation of raw data repository."""
    
    def get_file_categories(self) -> List[FileCategory]:
        """Get all available file categories from delivery catalog."""
        query = """
        SELECT FILE_CATEGORY 
        FROM test.dbo.DELIVERY_CATALOG d 
        WHERE GETDATE() > d.VALID_FROM AND GETDATE() < d.VALID_TO
        """
        
        try:
            results = db_manager.execute_query(query)
            return [FileCategory(category=row['FILE_CATEGORY']) for row in results]
        except Exception as e:
            logger.error(f"Failed to get file categories: {e}")
            raise
    
    def get_funds_by_catalog(self, catalog: str) -> List[Fund]:
        """Get all funds for a specific catalog."""
        query = f"SELECT DISTINCT FUND FROM test.dbo.{catalog}"
        
        try:
            results = db_manager.execute_query(query)
            return [Fund(fund=row['FUND'], catalog=catalog) for row in results]
        except Exception as e:
            logger.error(f"Failed to get funds for catalog {catalog}: {e}")
            raise
    
    def get_raw_data(self, request: RawDataRequest) -> List[RawDataRecord]:
        """Get raw data based on request parameters."""
        query = f"""
        SELECT * 
        FROM test.dbo.{request.catalog} c 
        WHERE c.fund = :fund 
        AND c.START_DATE BETWEEN :start AND :end
        """
        
        params = request.to_query_params()
        
        try:
            results = db_manager.execute_query(query, params)
            return [RawDataRecord(data=row) for row in results]
        except Exception as e:
            logger.error(f"Failed to get raw data: {e}")
            logger.error(f"Request: {request}")
            raise