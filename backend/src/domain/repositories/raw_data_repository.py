"""
Repository interface for raw data operations.
"""
from abc import ABC, abstractmethod
from typing import List
from ..entities.raw_data import FileCategory, Fund, RawDataRecord, RawDataRequest


class IRawDataRepository(ABC):
    """Interface for raw data repository operations."""
    
    @abstractmethod
    def get_file_categories(self) -> List[FileCategory]:
        """Get all available file categories."""
        pass
    
    @abstractmethod
    def get_funds_by_catalog(self, catalog: str) -> List[Fund]:
        """Get all funds for a specific catalog."""
        pass
    
    @abstractmethod
    def get_raw_data(self, request: RawDataRequest) -> List[RawDataRecord]:
        """Get raw data based on request parameters."""
        pass