"""
Fund column mappings for different table types.
"""
from typing import Dict, Optional

# Mapping of table names to their fund column names
FUND_COLUMN_MAPPINGS: Dict[str, str] = {
    'CITCO_ALLOC': 'fs_desc',
    'CITCO_CAPACT': 'CODE',
    'CITCO_PCAM_AEXBONDACCR': 'FUND',
    'CITCO_PCAM_DEBT_COUPON_PAYMENTS_SUMMARY': 'FUND',
    'CITCO_PCAM_DETAIL_TB': 'FUND',
    'CITCO_PCAM_DIV': 'FUND',
    'CITCO_PCAM_FX_CURVAL': 'FUND',
    'CITCO_PCAM_GL': 'FUND',
    'CITCO_PCAM_MONTHLYTB': 'FUND',
    'CITCO_PCAM_PORTFHOLD': 'FUND',
    'CITCO_PCAM_REALIZED_UNREALIZED_0': 'FUND_ABBREV',
    'CITCO_PCAM_UNSETTLED': 'FUND',
    'CITCO_PCAM_YETB': 'FUND',
    'CITCO_NAV': 'FUND_ABBREV',
    'CITCO_CAR': 'FS_DESC',
    'CITCO_SPOS_PCAM': 'SUBFUND',
    'CITCO_SECUR': 'CLIENT_NAME',
    'HAAS_DI_SUMMARY': 'FUND_NAME',
    'HAAS_DIRECT_TRANSACTIONS': 'FUND_NAME',
    'HAAS_FI_SUMMARY': 'FUND_NAME',
    'HAAS_FI_TRANSACTIONS': 'FUND_NAME',
    'HAAS_FUND_PRICE': 'FUND_NAME',
    'HAAS_GL_TRANSACTIONS': 'FUND_NAME',
    'HAAS_INVESTOR_DATA': 'FUND_NAME',
    'HAAS_INVESTOR_OPERATIONS': 'FUND_NAME',
    'HAAS_TRIAL_BALANCE': 'FUND_NAME',
}


def get_fund_column(catalog: str) -> Optional[str]:
    """
    Get the fund column name for a given catalog.
    
    Args:
        catalog: The table/catalog name
        
    Returns:
        The fund column name if mapped, None if not mapped
    """
    return FUND_COLUMN_MAPPINGS.get(catalog)


def has_fund_filtering(catalog: str) -> bool:
    """
    Check if the catalog supports fund filtering.
    
    Args:
        catalog: The table/catalog name
        
    Returns:
        True if fund filtering is supported, False otherwise
    """
    return catalog in FUND_COLUMN_MAPPINGS