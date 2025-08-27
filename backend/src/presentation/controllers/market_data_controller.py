"""
FastAPI controller for market data endpoints.
"""
from fastapi import APIRouter, HTTPException
from typing import Dict, Any
import logging

from ...application.services.market_data_service import MarketDataService
from ...application.dtos.data_dtos import MarketDataDownloadRequestDto

logger = logging.getLogger(__name__)

market_data_router = APIRouter(prefix='/api/market-data', tags=['market-data'])


@market_data_router.get('/securities')
async def get_securities():
    """Get securities for dropdown menu."""
    try:
        service = MarketDataService()
        securities = service.get_securities()
        
        # Convert to simple list for frontend dropdown
        security_list = [security.security for security in securities]
        
        return {
            'success': True,
            'data': security_list
        }
        
    except Exception as e:
        logger.error(f"Error getting securities: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@market_data_router.get('/fields/{security}')
async def get_fields(security: str):
    """Get fields for a specific security."""
    try:
        service = MarketDataService()
        fields = service.get_fields_by_security(security)
        
        # Convert to simple list for frontend dropdown
        field_list = [field.field for field in fields]
        
        return {
            'success': True,
            'data': field_list
        }
        
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
async def download_market_data(request_data: Dict[str, Any]):
    """Download market data based on filters."""
    try:
        if not request_data:
            raise HTTPException(
                status_code=400,
                detail={
                    'success': False,
                    'error': 'No data provided'
                }
            )
        
        # Validate required fields
        required_fields = ['security', 'field', 'start_date', 'end_date']
        for field in required_fields:
            if field not in request_data:
                raise HTTPException(
                    status_code=400,
                    detail={
                        'success': False,
                        'error': f'Missing required field: {field}'
                    }
                )
        
        # Create request DTO
        request_dto = MarketDataDownloadRequestDto(
            security=request_data['security'],
            field=request_data['field'],
            start_date=request_data['start_date'],
            end_date=request_data['end_date']
        )
        
        service = MarketDataService()
        
        # Check if batching is requested
        batch_size = request_data.get('batch_size', 1000)
        batch_id = request_data.get('batch_id', None)
        
        if batch_id is not None:
            # Return batched response
            result = service.download_market_data_batched(request_dto, batch_size, batch_id)
            return {
                'success': True,
                'batch_id': result.batch_id,
                'total_batches': result.total_batches,
                'has_more': result.has_more,
                'data': result.data
            }
        else:
            # Return all data
            records = service.download_market_data(request_dto)
            data_list = [record.data for record in records]
            
            # Get column names from first record to preserve order
            columns = []
            if data_list:
                columns = list(data_list[0].keys())
            
            return {
                'success': True,
                'count': len(data_list),
                'columns': columns,  # Preserve column order from database
                'data': data_list
            }
    
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