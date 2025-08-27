"""
Data Upload Controller for handling Excel data uploads and forwarding to NiFi.
"""
import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional

import requests
from fastapi import APIRouter, HTTPException

from src.infrastructure.config.app_config import AppConfig

logger = logging.getLogger(__name__)

# Create router
data_upload_router = APIRouter(prefix='/api/data-upload', tags=['data-upload'])


@data_upload_router.post('/upload')
async def upload_data(request_data: Dict[str, Any]):
    """
    Handle data upload from Excel Add-in.
    Process the data and forward to NiFi endpoint.
    """
    try:
        if not request_data:
            raise HTTPException(
                status_code=400,
                detail={
                    'success': False,
                    'error': 'No data provided'
                }
            )
        
        # Extract upload parameters
        data_type = request_data.get('dataType')
        skip_duplicate_check = request_data.get('skipDuplicateCheck', False)
        delivery_date = request_data.get('deliveryDate')
        data = request_data.get('data', [])
        
        # Validate required fields
        if not data_type:
            raise HTTPException(
                status_code=400,
                detail={
                    'success': False,
                    'error': 'Data type is required'
                }
            )
        
        if not data or not isinstance(data, list):
            raise HTTPException(
                status_code=400,
                detail={
                    'success': False,
                    'error': 'Data array is required and must be non-empty'
                }
            )
        
        # Log upload attempt
        logger.info(f"Processing data upload: type={data_type}, records={len(data)}")
        
        # Process and validate data
        processed_data = []
        for i, record in enumerate(data):
            if not isinstance(record, dict):
                raise HTTPException(
                    status_code=400,
                    detail={
                        'success': False,
                        'error': f'Invalid record format at index {i}'
                    }
                )
            
            # Add metadata to each record
            processed_record = {
                **record,
                '_upload_metadata': {
                    'upload_timestamp': datetime.utcnow().isoformat(),
                    'data_type': data_type,
                    'skip_duplicate_check': skip_duplicate_check,
                    'delivery_date': delivery_date,
                    'record_index': i
                }
            }
            processed_data.append(processed_record)
        
        # Prepare payload for NiFi
        nifi_payload = {
            'source': 'excel_addin',
            'upload_timestamp': datetime.utcnow().isoformat(),
            'data_type': data_type,
            'configuration': {
                'skip_duplicate_check': skip_duplicate_check,
                'delivery_date': delivery_date
            },
            'records': processed_data,
            'record_count': len(processed_data)
        }
        
        # Forward to NiFi endpoint
        nifi_endpoint = AppConfig.NIFI_ENDPOINT
        ssl_config = AppConfig.get_nifi_ssl_config()
        
        logger.info(f"Forwarding data to NiFi endpoint: {nifi_endpoint}")
        logger.debug(f"SSL configuration: verify={ssl_config.get('verify', 'default')}, "
                    f"client_cert={'configured' if ssl_config.get('cert') else 'not configured'}")
        
        try:
            response = requests.post(
                nifi_endpoint,
                json=nifi_payload,
                headers={
                    'Content-Type': 'application/json',
                    'X-Forwarded-From': 'excel-addin-backend'
                },
                timeout=30,
                **ssl_config  # Apply SSL configuration (verify, cert)
            )
            
            if response.status_code == 200 or response.status_code == 201:
                logger.info(f"Successfully forwarded {len(processed_data)} records to NiFi")
                return {
                    'success': True,
                    'message': f'Successfully uploaded {len(processed_data)} records',
                    'record_count': len(processed_data),
                    'data_type': data_type,
                    'nifi_response_status': response.status_code
                }
            else:
                logger.error(f"NiFi endpoint returned status {response.status_code}: {response.text}")
                raise HTTPException(
                    status_code=502,
                    detail={
                        'success': False,
                        'error': f'NiFi processing failed with status {response.status_code}',
                        'details': response.text[:500] if response.text else None
                    }
                )
                
        except requests.exceptions.SSLError as e:
            logger.error(f"SSL error when connecting to NiFi endpoint: {str(e)}")
            raise HTTPException(
                status_code=502,
                detail={
                    'success': False,
                    'error': 'SSL certificate verification failed when connecting to NiFi',
                    'details': 'Check certificate configuration in backend/certificates/ directory',
                    'ssl_error': str(e)
                }
            )
            
        except requests.exceptions.Timeout:
            logger.error("Timeout when connecting to NiFi endpoint")
            raise HTTPException(
                status_code=504,
                detail={
                    'success': False,
                    'error': 'Upload timeout - NiFi endpoint did not respond in time'
                }
            )
            
        except requests.exceptions.ConnectionError:
            logger.error("Connection error when connecting to NiFi endpoint")
            raise HTTPException(
                status_code=502,
                detail={
                    'success': False,
                    'error': 'Unable to connect to NiFi endpoint'
                }
            )
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Request error when connecting to NiFi: {str(e)}")
            raise HTTPException(
                status_code=502,
                detail={
                    'success': False,
                    'error': f'Request failed: {str(e)}'
                }
            )
    
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error in upload_data: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': 'Internal server error during upload processing'
            }
        )


@data_upload_router.get('/types')
async def get_upload_types():
    """
    Get available data upload types.
    """
    try:
        upload_types = [
            {
                'id': 'windmill_statistics',
                'name': 'Windmill Statistics',
                'description': 'Statistical data for windmill performance analysis'
            },
            {
                'id': 'financial_outperformance',
                'name': 'Financial Outperformance',
                'description': 'Financial performance comparison data'
            },
            {
                'id': 'excellence_accounting',
                'name': 'Excellence Accounting',
                'description': 'Accounting excellence metrics and data'
            }
        ]
        
        return {
            'success': True,
            'upload_types': upload_types
        }
    
    except Exception as e:
        logger.error(f"Error getting upload types: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': 'Failed to retrieve upload types'
            }
        )


@data_upload_router.get('/status/{upload_id}')
async def get_upload_status(upload_id: str):
    """
    Get the status of a specific upload.
    Note: This is a placeholder for future implementation with upload tracking.
    """
    try:
        # This would normally query a database for upload status
        # For now, return a simple response
        return {
            'success': True,
            'upload_id': upload_id,
            'status': 'completed',
            'message': 'Upload status tracking not yet implemented'
        }
    
    except Exception as e:
        logger.error(f"Error getting upload status: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail={
                'success': False,
                'error': 'Failed to retrieve upload status'
            }
        )