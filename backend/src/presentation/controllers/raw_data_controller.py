"""
FastAPI controller for raw data endpoints.
"""
from fastapi import APIRouter, HTTPException
from typing import Dict, Any
import logging

from ...application.services.raw_data_service import RawDataService
from ...application.dtos.data_dtos import RawDataDownloadRequestDto
from ...infrastructure.config.fund_mappings import has_fund_filtering
from ..models.request_models import RawDataDownloadRequest, DataResponse, BatchedDataResponse

logger = logging.getLogger(__name__)

raw_data_router = APIRouter(prefix='/api/raw-data', tags=['raw-data'])


@raw_data_router.get('/categories', response_model=DataResponse)
async def get_categories():
    """Get file categories for dropdown menu."""
    try:
        service = RawDataService()
        categories = service.get_file_categories()
        
        # Convert to simple list for frontend dropdown
        category_list = [category.category for category in categories]
        
        return DataResponse(
            success=True,
            data=category_list
        )
        
    except Exception as e:
        logger.error(f"Error getting categories: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@raw_data_router.get('/funds/{catalog}', response_model=DataResponse)
async def get_funds(catalog: str):
    """Get funds for a specific catalog."""
    try:
        # Check if fund filtering is available for this catalog
        if not has_fund_filtering(catalog):
            return DataResponse(
                success=True,
                data=[],
                fund_filtering_available=False
            )
        
        service = RawDataService()
        funds = service.get_funds_by_catalog(catalog)
        
        # Convert to simple list for frontend dropdown
        fund_list = [fund.fund for fund in funds]
        
        return DataResponse(
            success=True,
            data=fund_list,
            fund_filtering_available=True
        )
        
    except Exception as e:
        logger.error(f"Error getting funds for catalog {catalog}: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@raw_data_router.post('/download')
async def download_raw_data(request_data: RawDataDownloadRequest):
    """Download raw data based on filters."""
    try:
        # Check if fund is required for this catalog
        if has_fund_filtering(request_data.catalog):
            if not request_data.fund:
                raise HTTPException(
                    status_code=400,
                    detail={
                        'success': False,
                        'error': 'Fund is required for this catalog'
                    }
                )
        
        # Create request DTO - use empty string for fund if not available
        request_dto = RawDataDownloadRequestDto(
            catalog=request_data.catalog,
            fund=request_data.fund or '',
            start_date=request_data.start_date,
            end_date=request_data.end_date
        )
        
        service = RawDataService()
        
        if request_data.batch_id is not None:
            # Return batched response
            result = service.download_raw_data_batched(request_dto, request_data.batch_size, request_data.batch_id)
            return BatchedDataResponse(
                success=True,
                batch_id=result.batch_id,
                total_batches=result.total_batches,
                has_more=result.has_more,
                data=result.data
            )
        else:
            # Return all data
            records = service.download_raw_data(request_dto)
            data_list = [record.data for record in records]
            
            # Get column names from first record to preserve order
            columns = []
            if data_list:
                columns = list(data_list[0].keys())
            
            return DataResponse(
                success=True,
                count=len(data_list),
                columns=columns,  # Preserve column order from database
                data=data_list
            )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error downloading raw data: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )