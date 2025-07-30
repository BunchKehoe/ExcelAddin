"""
Mock repository implementations for testing without database connectivity.
"""
from typing import List
import logging

from ...domain.repositories.raw_data_repository import IRawDataRepository
from ...domain.repositories.market_data_repository import IMarketDataRepository
from ...domain.entities.raw_data import FileCategory, Fund, RawDataRecord, RawDataRequest
from ...domain.entities.market_data import Security, DataField, MarketDataRecord, MarketDataRequest

logger = logging.getLogger(__name__)


class MockRawDataRepository(IRawDataRepository):
    """Mock implementation of raw data repository for testing."""
    
    def get_file_categories(self) -> List[FileCategory]:
        """Get mock file categories."""
        return [
            FileCategory(category="SAMPLE_CATEGORY_1"),
            FileCategory(category="SAMPLE_CATEGORY_2"),
            FileCategory(category="NAV_DATA"),
            FileCategory(category="BALANCE_SHEET"),
            FileCategory(category="TRANSACTIONS")
        ]
    
    def get_funds_by_catalog(self, catalog: str) -> List[Fund]:
        """Get mock funds for a catalog."""
        return [
            Fund(fund="GLOBAL_EQUITY_FUND", catalog=catalog),
            Fund(fund="FIXED_INCOME_FUND", catalog=catalog),
            Fund(fund="EMERGING_MARKETS_FUND", catalog=catalog),
            Fund(fund="TECHNOLOGY_FUND", catalog=catalog),
            Fund(fund="REAL_ESTATE_FUND", catalog=catalog)
        ]
    
    def get_raw_data(self, request: RawDataRequest) -> List[RawDataRecord]:
        """Get mock raw data."""
        # Generate mock data based on date range
        import datetime
        
        data = []
        current_date = request.start_date
        value = 100.0
        
        while current_date <= request.end_date:
            # Add some variance to the data
            import random
            value += random.uniform(-5, 5)
            
            record_data = {
                'date': current_date.strftime('%Y-%m-%d'),
                'fund': request.fund,
                'catalog': request.catalog,
                'value': round(value, 2),
                'nav': round(value * 1.1, 2),
                'shares': random.randint(1000, 10000),
                'currency': 'USD'
            }
            
            data.append(RawDataRecord(data=record_data))
            current_date += datetime.timedelta(days=1)
        
        return data[:50]  # Limit to 50 records for demo


class MockMarketDataRepository(IMarketDataRepository):
    """Mock implementation of market data repository for testing."""
    
    def get_securities(self) -> List[Security]:
        """Get mock securities."""
        return [
            Security(security="AAPL US Equity"),
            Security(security="MSFT US Equity"),
            Security(security="GOOGL US Equity"),
            Security(security="TSLA US Equity"),
            Security(security="AMZN US Equity")
        ]
    
    def get_fields_by_security(self, security: str) -> List[DataField]:
        """Get mock fields for a security."""
        return [
            DataField(field="PX_LAST", security=security),
            DataField(field="PX_OPEN", security=security),
            DataField(field="PX_HIGH", security=security),
            DataField(field="PX_LOW", security=security),
            DataField(field="PX_VOLUME", security=security),
            DataField(field="MARKET_CAP", security=security)
        ]
    
    def get_market_data(self, request: MarketDataRequest) -> List[MarketDataRecord]:
        """Get mock market data."""
        import datetime
        import random
        
        data = []
        current_date = request.start_date
        base_price = 150.0 if "AAPL" in request.security else random.uniform(50, 500)
        
        while current_date <= request.end_date:
            # Generate different values based on field type
            if request.field == "PX_LAST":
                value = base_price + random.uniform(-10, 10)
            elif request.field == "PX_VOLUME":
                value = random.randint(1000000, 50000000)
            elif request.field == "MARKET_CAP":
                value = random.randint(100000000000, 3000000000000)
            else:
                value = base_price + random.uniform(-15, 15)
            
            record = MarketDataRecord(
                security=request.security,
                field=request.field,
                date=current_date,
                value=round(value, 2)
            )
            
            data.append(record)
            current_date += datetime.timedelta(days=1)
        
        return data[:50]  # Limit to 50 records for demo