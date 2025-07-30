"""
SQL Server implementation of raw data repository.
"""
from typing import List
import logging

from ...domain.repositories.raw_data_repository import IRawDataRepository
from ...domain.entities.raw_data import FileCategory, Fund, RawDataRecord, RawDataRequest
from ..database.db_manager import db_manager
from ..config.fund_mappings import get_fund_column, has_fund_filtering

logger = logging.getLogger(__name__)


class SqlRawDataRepository(IRawDataRepository):
    """SQL Server implementation of raw data repository."""
    
    def get_file_categories(self) -> List[FileCategory]:
        """Get all available file categories from information schema tables."""
        query = """
        SELECT t.TABLE_NAME
        FROM INFORMATION_SCHEMA.TABLES t
        JOIN test.dbo.DATA_PROVIDERS d
          ON t.TABLE_NAME LIKE d.PROVIDER_NAME + '%'
         WHERE TABLE_SCHEMA = 'dbo' 
         AND TABLE_TYPE = 'BASE TABLE'
         AND TABLE_NAME NOT IN ('BLOOMBERG_ODD_MONTHLY')
         ORDER BY TABLE_NAME asc
        """
        
        try:
            results = db_manager.execute_query(query)
            return [FileCategory(category=row['TABLE_NAME']) for row in results]
        except Exception as e:
            logger.error(f"Failed to get file categories: {e}")
            raise
    
    def get_funds_by_catalog(self, catalog: str) -> List[Fund]:
        """Get all funds for a specific catalog."""
        fund_column = get_fund_column(catalog)
        
        if not fund_column:
            # Return empty list if no fund filtering is available for this catalog
            return []
        
        query = f"SELECT DISTINCT {fund_column} as FUND FROM test.dbo.{catalog}"
        
        try:
            results = db_manager.execute_query(query)
            return [Fund(fund=row['FUND'], catalog=catalog) for row in results if row['FUND'] is not None]
        except Exception as e:
            logger.error(f"Failed to get funds for catalog {catalog}: {e}")
            raise
    
    def get_raw_data(self, request: RawDataRequest) -> List[RawDataRecord]:
        """Get raw data based on request parameters."""
        fund_column = get_fund_column(request.catalog)
        
        if fund_column:
            # Query with fund filtering
            query = f"""
            SELECT c.* 
            FROM test.dbo.{request.catalog} c 
            JOIN test.dbo.DELIVERY d on d.DELIVERY_ID = c.DELIVERY_ID
            WHERE c.{fund_column} = :fund 
            AND CAST(c.LOAD_TS AS DATE) BETWEEN :start AND :end
            """
            params = request.to_query_params()
        else:
            # Query without fund filtering for tables not in mapping
            query = f"""
            SELECT c.* 
            FROM test.dbo.{request.catalog} c 
            JOIN test.dbo.DELIVERY d on d.DELIVERY_ID = c.DELIVERY_ID
            WHERE CAST(c.LOAD_TS AS DATE) BETWEEN :start AND :end
            """
            # Remove fund parameter for tables without fund filtering
            params = {
                'start': request.start_date,
                'end': request.end_date
            }
        
        try:
            results = db_manager.execute_query(query, params)
            return [RawDataRecord(data=row) for row in results]
        except Exception as e:
            logger.error(f"Failed to get raw data: {e}")
            logger.error(f"Request: {request}")
            raise