"""
FastAPI controller for market data endpoints.
"""
from fastapi import APIRouter, HTTPException
from typing import Dict, Any
import logging

from ...application.services.market_data_service import MarketDataService
from ...application.dtos.data_dtos import MarketDataDownloadRequestDto
from ..models.request_models import MarketDataDownloadRequest, DataResponse, BatchedDataResponse

logger = logging.getLogger(__name__)

market_data_router = APIRouter(prefix='/api/market-data', tags=['market-data'])


@market_data_router.get('/securities', response_model=DataResponse)
async def get_securities():
    """Get securities for dropdown menu."""
    try:
        service = MarketDataService()
        securities = service.get_securities()
        
        # Convert to simple list for frontend dropdown
        security_list = [security.security for security in securities]
        
        return DataResponse(
            success=True,
            data=security_list
        )
        
    except Exception as e:
        logger.error(f"Error getting securities: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@market_data_router.get('/fields/{security}', response_model=DataResponse)
async def get_fields(security: str):
    """Get fields for a specific security."""
    try:
        service = MarketDataService()
        fields = service.get_fields_by_security(security)
        
        # Convert to simple list for frontend dropdown
        field_list = [field.field for field in fields]
        
        return DataResponse(
            success=True,
            data=field_list
        )
        
    except Exception as e:
        logger.error(f"Error getting fields for security {security}: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@market_data_router.post('/download')
async def download_market_data(request_data: MarketDataDownloadRequest):
    """Download market data based on filters."""
    try:
        # Create request DTO
        request_dto = MarketDataDownloadRequestDto(
            security=request_data.security,
            field=request_data.field,
            start_date=request_data.start_date,
            end_date=request_data.end_date
        )
        
        service = MarketDataService()
        
        if request_data.batch_id is not None:
            # Return batched response
            result = service.download_market_data_batched(request_dto, request_data.batch_size, request_data.batch_id)
            return BatchedDataResponse(
                success=True,
                batch_id=result.batch_id,
                total_batches=result.total_batches,
                has_more=result.has_more,
                data=result.data
            )
        else:
            # Return all data
            records = service.download_market_data(request_dto)
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
        logger.error(f"Error downloading market data: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )