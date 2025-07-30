"""
Application service for raw data operations.
"""
from typing import List
from datetime import datetime
import logging

from ..dtos.data_dtos import (
    FileCategoryDto, FundDto, RawDataDownloadRequestDto, 
    DataRecordDto, BatchedDataResponseDto
)
from ...domain.repositories.raw_data_repository import IRawDataRepository
from ...domain.entities.raw_data import RawDataRequest

logger = logging.getLogger(__name__)


class RawDataService:
    """Application service for raw data operations."""
    
    def __init__(self, repository: IRawDataRepository = None):
        if repository:
            self._repository = repository
        else:
            # Try SQL repository first, fall back to mock if connection fails
            try:
                from ...infrastructure.database.sql_raw_data_repository import SqlRawDataRepository
                self._repository = SqlRawDataRepository()
                # Test the connection
                self._repository.get_file_categories()
                logger.info("Using SQL Server repository")
            except Exception as e:
                logger.warning(f"Database connection failed, using mock repository: {e}")
                from ...infrastructure.database.mock_repositories import MockRawDataRepository
                self._repository = MockRawDataRepository()
    
    def get_file_categories(self) -> List[FileCategoryDto]:
        """Get all available file categories for dropdown."""
        try:
            categories = self._repository.get_file_categories()
            return [FileCategoryDto(category=cat.category) for cat in categories]
        except Exception as e:
            logger.error(f"Failed to get file categories: {e}")
            raise
    
    def get_funds_by_catalog(self, catalog: str) -> List[FundDto]:
        """Get all funds for a specific catalog for dropdown."""
        try:
            funds = self._repository.get_funds_by_catalog(catalog)
            return [FundDto(fund=fund.fund, catalog=fund.catalog) for fund in funds]
        except Exception as e:
            logger.error(f"Failed to get funds for catalog {catalog}: {e}")
            raise
    
    def download_raw_data(self, request_dto: RawDataDownloadRequestDto, batch_size: int = 1000) -> List[DataRecordDto]:
        """Download raw data with optional batching."""
        try:
            # Convert DTO to domain entity
            request = RawDataRequest(
                catalog=request_dto.catalog,
                fund=request_dto.fund,
                start_date=datetime.fromisoformat(request_dto.start_date),
                end_date=datetime.fromisoformat(request_dto.end_date)
            )
            
            # Get data from repository
            raw_data = self._repository.get_raw_data(request)
            
            # Convert to DTOs
            data_records = [DataRecordDto(data=record.to_dict()) for record in raw_data]
            
            logger.info(f"Retrieved {len(data_records)} raw data records")
            return data_records
            
        except Exception as e:
            logger.error(f"Failed to download raw data: {e}")
            raise
    
    def download_raw_data_batched(self, request_dto: RawDataDownloadRequestDto, 
                                 batch_size: int = 1000, batch_id: int = 0) -> BatchedDataResponseDto:
        """Download raw data in batches for large datasets."""
        try:
            # Get all data first (in a real implementation, this could be optimized with OFFSET/LIMIT)
            all_data = self.download_raw_data(request_dto, batch_size)
            
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
            logger.error(f"Failed to download batched raw data: {e}")
            raise