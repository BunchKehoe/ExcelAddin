"""
Domain entities for market data management.
"""
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional, Any, Dict


@dataclass
class Security:
    """Represents a security in the market data."""
    security: str


@dataclass
class DataField:
    """Represents a data field for a specific security."""
    field: str
    security: str


@dataclass
class MarketDataRecord:
    """Represents a market data record."""
    security: str
    field: str
    date: datetime
    value: Any
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return {
            'security': self.security,
            'field': self.field,
            'date': self.date.isoformat() if isinstance(self.date, datetime) else self.date,
            'value': self.value
        }


@dataclass
class MarketDataRequest:
    """Request parameters for market data download."""
    security: str
    field: str
    start_date: datetime
    end_date: datetime
    
    def to_query_params(self) -> Dict[str, Any]:
        """Convert to database query parameters."""
        return {
            'security': self.security,
            'field': self.field,
            'start': self.start_date,
            'end': self.end_date
        }