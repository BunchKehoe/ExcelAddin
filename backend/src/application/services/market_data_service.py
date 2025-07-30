"""
Application service for market data operations.
"""
from typing import List
from datetime import datetime
import logging

from ..dtos.data_dtos import (
    SecurityDto, DataFieldDto, MarketDataDownloadRequestDto, 
    DataRecordDto, BatchedDataResponseDto
)
from ...domain.repositories.market_data_repository import IMarketDataRepository
from ...domain.entities.market_data import MarketDataRequest

logger = logging.getLogger(__name__)


class MarketDataService:
    """Application service for market data operations."""
    
    def __init__(self, repository: IMarketDataRepository = None):
        if repository:
            self._repository = repository
        else:
            # Try SQL repository first, fall back to mock if connection fails
            try:
                from ...infrastructure.database.sql_market_data_repository import SqlMarketDataRepository
                self._repository = SqlMarketDataRepository()
                # Test the connection
                self._repository.get_securities()
                logger.info("Using SQL Server repository for market data")
            except Exception as e:
                logger.warning(f"Database connection failed, using mock repository for market data: {e}")
                from ...infrastructure.database.mock_repositories import MockMarketDataRepository
                self._repository = MockMarketDataRepository()
    
    def get_securities(self) -> List[SecurityDto]:
        """Get all available securities for dropdown."""
        try:
            securities = self._repository.get_securities()
            return [SecurityDto(security=sec.security) for sec in securities]
        except Exception as e:
            logger.error(f"Failed to get securities: {e}")
            raise
    
    def get_fields_by_security(self, security: str) -> List[DataFieldDto]:
        """Get all fields for a specific security for dropdown."""
        try:
            fields = self._repository.get_fields_by_security(security)
            return [DataFieldDto(field=field.field, security=field.security) for field in fields]
        except Exception as e:
            logger.error(f"Failed to get fields for security {security}: {e}")
            raise
    
    def download_market_data(self, request_dto: MarketDataDownloadRequestDto) -> List[DataRecordDto]:
        """Download market data."""
        try:
            # Convert DTO to domain entity
            request = MarketDataRequest(
                security=request_dto.security,
                field=request_dto.field,
                start_date=datetime.fromisoformat(request_dto.start_date),
                end_date=datetime.fromisoformat(request_dto.end_date)
            )
            
            # Get data from repository
            market_data = self._repository.get_market_data(request)
            
            # Convert to DTOs
            data_records = [DataRecordDto(data=record.to_dict()) for record in market_data]
            
            logger.info(f"Retrieved {len(data_records)} market data records")
            return data_records
            
        except Exception as e:
            logger.error(f"Failed to download market data: {e}")
            raise
    
    def download_market_data_batched(self, request_dto: MarketDataDownloadRequestDto, 
                                   batch_size: int = 1000, batch_id: int = 0) -> BatchedDataResponseDto:
        """Download market data in batches for large datasets."""
        try:
            # Get all data first
            all_data = self.download_market_data(request_dto)
            
            # Calculate batch boundaries
            start_idx = batch_id * batch_size
            end_idx = start_idx + batch_size
            total_records = len(all_data)
            total_batches = (total_records + batch_size - 1) // batch_size
            
            # Get batch data
            batch_data = all_data[start_idx:end_idx]
            has_more = end_idx < total_records
            
            return BatchedDataResponseDto(
                batch_id=batch_id,
                total_batches=total_batches,
                data=[record.data for record in batch_data],
                has_more=has_more
            )
            
        except Exception as e:
            logger.error(f"Failed to download batched market data: {e}")
            raise