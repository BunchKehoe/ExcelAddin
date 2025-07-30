"""
Repository interface for market data operations.
"""
from abc import ABC, abstractmethod
from typing import List
from ..entities.market_data import Security, DataField, MarketDataRecord, MarketDataRequest


class IMarketDataRepository(ABC):
    """Interface for market data repository operations."""
    
    @abstractmethod
    def get_securities(self) -> List[Security]:
        """Get all available securities."""
        pass
    
    @abstractmethod
    def get_fields_by_security(self, security: str) -> List[DataField]:
        """Get all fields for a specific security."""
        pass
    
    @abstractmethod
    def get_market_data(self, request: MarketDataRequest) -> List[MarketDataRecord]:
        """Get market data based on request parameters."""
        pass