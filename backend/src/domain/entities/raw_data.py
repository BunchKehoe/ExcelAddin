"""
Domain entities for raw data management.
"""
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional, Any, Dict


@dataclass
class FileCategory:
    """Represents a file category from the delivery catalog."""
    category: str


@dataclass
class Fund:
    """Represents a fund within a specific catalog."""
    fund: str
    catalog: str


@dataclass
class RawDataRecord:
    """Represents a record from raw data tables."""
    data: Dict[str, Any]  # Dynamic data structure
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for JSON serialization."""
        return self.data


@dataclass
class RawDataRequest:
    """Request parameters for raw data download."""
    catalog: str
    fund: str
    start_date: datetime
    end_date: datetime
    
    def to_query_params(self) -> Dict[str, Any]:
        """Convert to database query parameters."""
        return {
            'fund': self.fund,
            'start': self.start_date,
            'end': self.end_date
        }