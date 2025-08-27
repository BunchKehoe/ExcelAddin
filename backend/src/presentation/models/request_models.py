"""
Pydantic models for FastAPI request/response validation.
"""
from pydantic import BaseModel, Field
from typing import List, Dict, Any, Optional


class RawDataDownloadRequest(BaseModel):
    """Request model for raw data download."""
    catalog: str = Field(..., description="Data catalog name")
    fund: Optional[str] = Field(None, description="Fund name (optional)")
    start_date: str = Field(..., description="Start date in ISO format")
    end_date: str = Field(..., description="End date in ISO format")
    batch_size: Optional[int] = Field(1000, description="Batch size for paginated requests")
    batch_id: Optional[int] = Field(None, description="Batch ID for pagination")


class MarketDataDownloadRequest(BaseModel):
    """Request model for market data download."""
    security: str = Field(..., description="Security identifier")
    field: str = Field(..., description="Data field name")
    start_date: str = Field(..., description="Start date in ISO format")
    end_date: str = Field(..., description="End date in ISO format")
    batch_size: Optional[int] = Field(1000, description="Batch size for paginated requests")
    batch_id: Optional[int] = Field(None, description="Batch ID for pagination")


class DataUploadRequest(BaseModel):
    """Request model for data upload."""
    dataType: str = Field(..., description="Type of data being uploaded")
    skipDuplicateCheck: Optional[bool] = Field(False, description="Whether to skip duplicate checking")
    deliveryDate: Optional[str] = Field(None, description="Delivery date for the data")
    data: List[Dict[str, Any]] = Field(..., description="Array of data records to upload")


class StandardResponse(BaseModel):
    """Standard response model."""
    success: bool = Field(..., description="Whether the operation was successful")
    message: Optional[str] = Field(None, description="Response message")
    error: Optional[str] = Field(None, description="Error message if operation failed")
    
    class Config:
        extra = "allow"  # Allow extra fields for dynamic responses


class DataResponse(StandardResponse):
    """Response model for data endpoints."""
    data: Optional[List[Any]] = Field(None, description="Response data")
    count: Optional[int] = Field(None, description="Number of records returned")
    columns: Optional[List[str]] = Field(None, description="Column names for tabular data")
    fund_filtering_available: Optional[bool] = Field(None, description="Whether fund filtering is available")


class BatchedDataResponse(StandardResponse):
    """Response model for batched data endpoints."""
    batch_id: int = Field(..., description="Current batch ID")
    total_batches: int = Field(..., description="Total number of batches")
    has_more: bool = Field(..., description="Whether more batches are available")
    data: List[Dict[str, Any]] = Field(..., description="Data for this batch")


class UploadTypesResponse(StandardResponse):
    """Response model for upload types endpoint."""
    upload_types: List[Dict[str, str]] = Field(..., description="Available upload types")