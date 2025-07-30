"""
Data Transfer Objects for API communication.
"""
from dataclasses import dataclass
from datetime import datetime
from typing import List, Any, Dict, Optional


@dataclass
class FileCategoryDto:
    """DTO for file category."""
    category: str


@dataclass
class FundDto:
    """DTO for fund information."""
    fund: str
    catalog: str


@dataclass
class RawDataDownloadRequestDto:
    """DTO for raw data download request."""
    catalog: str
    fund: str
    start_date: str  # ISO format date string
    end_date: str    # ISO format date string


@dataclass
class SecurityDto:
    """DTO for security information."""
    security: str


@dataclass
class DataFieldDto:
    """DTO for data field information."""
    field: str
    security: str


@dataclass
class MarketDataDownloadRequestDto:
    """DTO for market data download request."""
    security: str
    field: str
    start_date: str  # ISO format date string
    end_date: str    # ISO format date string


@dataclass
class DataRecordDto:
    """DTO for generic data record."""
    data: Dict[str, Any]


@dataclass
class BatchedDataResponseDto:
    """DTO for batched data response."""
    batch_id: int
    total_batches: int
    data: List[Dict[str, Any]]
    has_more: bool