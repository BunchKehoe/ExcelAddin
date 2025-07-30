"""
Flask controller for raw data endpoints.
"""
from flask import Blueprint, request, jsonify
from typing import Dict, Any
import logging

from ...application.services.raw_data_service import RawDataService
from ...application.dtos.data_dtos import RawDataDownloadRequestDto
from ...infrastructure.config.fund_mappings import has_fund_filtering

logger = logging.getLogger(__name__)

raw_data_bp = Blueprint('raw_data', __name__, url_prefix='/api/raw-data')


@raw_data_bp.route('/categories', methods=['GET'])
def get_categories():
    """Get file categories for dropdown menu."""
    try:
        service = RawDataService()
        categories = service.get_file_categories()
        
        # Convert to simple list for frontend dropdown
        category_list = [category.category for category in categories]
        
        return jsonify({
            'success': True,
            'data': category_list
        })
        
    except Exception as e:
        logger.error(f"Error getting categories: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@raw_data_bp.route('/funds/<string:catalog>', methods=['GET'])
def get_funds(catalog: str):
    """Get funds for a specific catalog."""
    try:
        # Check if fund filtering is available for this catalog
        if not has_fund_filtering(catalog):
            return jsonify({
                'success': True,
                'data': [],
                'fund_filtering_available': False
            })
        
        service = RawDataService()
        funds = service.get_funds_by_catalog(catalog)
        
        # Convert to simple list for frontend dropdown
        fund_list = [fund.fund for fund in funds]
        
        return jsonify({
            'success': True,
            'data': fund_list,
            'fund_filtering_available': True
        })
        
    except Exception as e:
        logger.error(f"Error getting funds for catalog {catalog}: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@raw_data_bp.route('/download', methods=['POST'])
def download_raw_data():
    """Download raw data based on filters."""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                'success': False,
                'error': 'No data provided'
            }), 400
        
        # Check required fields
        required_fields = ['catalog', 'start_date', 'end_date']
        for field in required_fields:
            if field not in data:
                return jsonify({
                    'success': False,
                    'error': f'Missing required field: {field}'
                }), 400
        
        # Check if fund is required for this catalog
        if has_fund_filtering(data['catalog']):
            if 'fund' not in data or not data['fund']:
                return jsonify({
                    'success': False,
                    'error': 'Fund is required for this catalog'
                }), 400
        
        # Create request DTO - use empty string for fund if not available
        request_dto = RawDataDownloadRequestDto(
            catalog=data['catalog'],
            fund=data.get('fund', ''),
            start_date=data['start_date'],
            end_date=data['end_date']
        )
        
        service = RawDataService()
        
        # Check if batching is requested
        batch_size = data.get('batch_size', 1000)
        batch_id = data.get('batch_id', None)
        
        if batch_id is not None:
            # Return batched response
            result = service.download_raw_data_batched(request_dto, batch_size, batch_id)
            return jsonify({
                'success': True,
                'batch_id': result.batch_id,
                'total_batches': result.total_batches,
                'has_more': result.has_more,
                'data': result.data
            })
        else:
            # Return all data
            records = service.download_raw_data(request_dto)
            data_list = [record.data for record in records]
            
            return jsonify({
                'success': True,
                'count': len(data_list),
                'data': data_list
            })
    
    except Exception as e:
        logger.error(f"Error downloading raw data: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@raw_data_bp.errorhandler(404)
def not_found(error):
    """Handle 404 errors."""
    return jsonify({
        'success': False,
        'error': 'Endpoint not found'
    }), 404


@raw_data_bp.errorhandler(500)
def internal_error(error):
    """Handle 500 errors."""
    return jsonify({
        'success': False,
        'error': 'Internal server error'
    }), 500