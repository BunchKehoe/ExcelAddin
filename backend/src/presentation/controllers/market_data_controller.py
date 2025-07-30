"""
Flask controller for market data endpoints.
"""
from flask import Blueprint, request, jsonify
from typing import Dict, Any
import logging

from ...application.services.market_data_service import MarketDataService
from ...application.dtos.data_dtos import MarketDataDownloadRequestDto

logger = logging.getLogger(__name__)

market_data_bp = Blueprint('market_data', __name__, url_prefix='/api/market-data')


@market_data_bp.route('/securities', methods=['GET'])
def get_securities():
    """Get securities for dropdown menu."""
    try:
        service = MarketDataService()
        securities = service.get_securities()
        
        # Convert to simple list for frontend dropdown
        security_list = [security.security for security in securities]
        
        return jsonify({
            'success': True,
            'data': security_list
        })
        
    except Exception as e:
        logger.error(f"Error getting securities: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@market_data_bp.route('/fields/<string:security>', methods=['GET'])
def get_fields(security: str):
    """Get fields for a specific security."""
    try:
        service = MarketDataService()
        fields = service.get_fields_by_security(security)
        
        # Convert to simple list for frontend dropdown
        field_list = [field.field for field in fields]
        
        return jsonify({
            'success': True,
            'data': field_list
        })
        
    except Exception as e:
        logger.error(f"Error getting fields for security {security}: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@market_data_bp.route('/download', methods=['POST'])
def download_market_data():
    """Download market data based on filters."""
    try:
        data = request.get_json()
        
        if not data:
            return jsonify({
                'success': False,
                'error': 'No data provided'
            }), 400
        
        # Validate required fields
        required_fields = ['security', 'field', 'start_date', 'end_date']
        for field in required_fields:
            if field not in data:
                return jsonify({
                    'success': False,
                    'error': f'Missing required field: {field}'
                }), 400
        
        # Create request DTO
        request_dto = MarketDataDownloadRequestDto(
            security=data['security'],
            field=data['field'],
            start_date=data['start_date'],
            end_date=data['end_date']
        )
        
        service = MarketDataService()
        
        # Check if batching is requested
        batch_size = data.get('batch_size', 1000)
        batch_id = data.get('batch_id', None)
        
        if batch_id is not None:
            # Return batched response
            result = service.download_market_data_batched(request_dto, batch_size, batch_id)
            return jsonify({
                'success': True,
                'batch_id': result.batch_id,
                'total_batches': result.total_batches,
                'has_more': result.has_more,
                'data': result.data
            })
        else:
            # Return all data
            records = service.download_market_data(request_dto)
            data_list = [record.data for record in records]
            
            # Get column names from first record to preserve order
            columns = []
            if data_list:
                columns = list(data_list[0].keys())
            
            return jsonify({
                'success': True,
                'count': len(data_list),
                'columns': columns,  # Preserve column order from database
                'data': data_list
            })
    
    except Exception as e:
        logger.error(f"Error downloading market data: {e}")
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@market_data_bp.errorhandler(404)
def not_found(error):
    """Handle 404 errors."""
    return jsonify({
        'success': False,
        'error': 'Endpoint not found'
    }), 404


@market_data_bp.errorhandler(500)
def internal_error(error):
    """Handle 500 errors."""
    return jsonify({
        'success': False,
        'error': 'Internal server error'
    }), 500