"""
FastAPI controller for raw data endpoints.
"""
from fastapi import APIRouter, HTTPException
from typing import Dict, Any
import logging

from ...application.services.raw_data_service import RawDataService
from ...application.dtos.data_dtos import RawDataDownloadRequestDto
from ...infrastructure.config.fund_mappings import has_fund_filtering

logger = logging.getLogger(__name__)

raw_data_router = APIRouter(prefix='/api/raw-data', tags=['raw-data'])


@raw_data_router.get('/categories')
async def get_categories():
    """Get file categories for dropdown menu."""
    try:
        service = RawDataService()
        categories = service.get_file_categories()
        
        # Convert to simple list for frontend dropdown
        category_list = [category.category for category in categories]
        
        return {
            'success': True,
            'data': category_list
        }
        
    except Exception as e:
        logger.error(f"Error getting categories: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )


@raw_data_router.get('/funds/{catalog}')
async def get_funds(catalog: str):
    """Get funds for a specific catalog."""
    try:
        # Check if fund filtering is available for this catalog
        if not has_fund_filtering(catalog):
            return {
                'success': True,
                'data': [],
                'fund_filtering_available': False
            }
        
        service = RawDataService()
        funds = service.get_funds_by_catalog(catalog)
        
        # Convert to simple list for frontend dropdown
        fund_list = [fund.fund for fund in funds]
        
        return {
            'success': True,
            'data': fund_list,
            'fund_filtering_available': True
        }
        
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
async def download_raw_data(request_data: Dict[str, Any]):
    """Download raw data based on filters."""
    try:
        if not request_data:
            raise HTTPException(
                status_code=400,
                detail={
                    'success': False,
                    'error': 'No data provided'
                }
            )
        
        # Check required fields
        required_fields = ['catalog', 'start_date', 'end_date']
        for field in required_fields:
            if field not in request_data:
                raise HTTPException(
                    status_code=400,
                    detail={
                        'success': False,
                        'error': f'Missing required field: {field}'
                    }
                )
        
        # Check if fund is required for this catalog
        if has_fund_filtering(request_data['catalog']):
            if 'fund' not in request_data or not request_data['fund']:
                raise HTTPException(
                    status_code=400,
                    detail={
                        'success': False,
                        'error': 'Fund is required for this catalog'
                    }
                )
        
        # Create request DTO - use empty string for fund if not available
        request_dto = RawDataDownloadRequestDto(
            catalog=request_data['catalog'],
            fund=request_data.get('fund', ''),
            start_date=request_data['start_date'],
            end_date=request_data['end_date']
        )
        
        service = RawDataService()
        
        # Check if batching is requested
        batch_size = request_data.get('batch_size', 1000)
        batch_id = request_data.get('batch_id', None)
        
        if batch_id is not None:
            # Return batched response
            result = service.download_raw_data_batched(request_dto, batch_size, batch_id)
            return {
                'success': True,
                'batch_id': result.batch_id,
                'total_batches': result.total_batches,
                'has_more': result.has_more,
                'data': result.data
            }
        else:
            # Return all data
            records = service.download_raw_data(request_dto)
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
        logger.error(f"Error downloading raw data: {e}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': str(e)
            }
        )