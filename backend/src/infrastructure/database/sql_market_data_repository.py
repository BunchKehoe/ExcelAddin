"""
SQL Server implementation of market data repository.
"""
from typing import List
from datetime import datetime
import logging

from ...domain.repositories.market_data_repository import IMarketDataRepository
from ...domain.entities.market_data import Security, DataField, MarketDataRecord, MarketDataRequest
from ..database.db_manager import db_manager

logger = logging.getLogger(__name__)


class SqlMarketDataRepository(IMarketDataRepository):
    """SQL Server implementation of market data repository."""
    
    def get_securities(self) -> List[Security]:
        """Get all available securities from Bloomberg data."""
        query = "SELECT DISTINCT b.security FROM BLOOMBERG_ODD_MONTHLY b"
        
        try:
            results = db_manager.execute_query(query)
            return [Security(security=row['security']) for row in results]
        except Exception as e:
            logger.error(f"Failed to get securities: {e}")
            raise
    
    def get_fields_by_security(self, security: str) -> List[DataField]:
        """Get all fields for a specific security."""
        query = """
        SELECT DISTINCT b.field 
        FROM BLOOMBERG_ODD_MONTHLY b 
        WHERE b.security = :security
        """
        
        params = {'security': security}
        
        try:
            results = db_manager.execute_query(query, params)
            return [DataField(field=row['field'], security=security) for row in results]
        except Exception as e:
            logger.error(f"Failed to get fields for security {security}: {e}")
            raise
    
    def get_market_data(self, request: MarketDataRequest) -> List[MarketDataRecord]:
        """Get market data based on request parameters."""
        query = """
        SELECT * 
        FROM BLOOMBERG_ODD_MONTHLY b 
        WHERE b.security = :security 
        AND b.field = :field 
        AND b.date BETWEEN :start AND :end
        """
        
        params = request.to_query_params()
        
        try:
            results = db_manager.execute_query(query, params)
            market_data = []
            
            for row in results:
                # Convert date if it's not already a datetime object
                date_value = row.get('date')
                if isinstance(date_value, str):
                    try:
                        date_value = datetime.fromisoformat(date_value)
                    except ValueError:
                        # Handle different date formats if needed
                        pass
                
                record = MarketDataRecord(
                    security=row.get('security'),
                    field=row.get('field'),
                    date=date_value,
                    value=row.get('value')  # Assuming there's a value column
                )
                market_data.append(record)
            
            return market_data
            
        except Exception as e:
            logger.error(f"Failed to get market data: {e}")
            logger.error(f"Request: {request}")
            raise