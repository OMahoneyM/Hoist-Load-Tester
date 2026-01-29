"""
FlexERP API Integration Module

This module handles barcode parsing and API communication with the FlexERP system
to fetch inventory data for hoist load testing.

Workflow:
1. Accept barcode input in format: ^12345 (where 12345 is the inventory ID)
2. Extract numeric inventory ID from barcode
3. Make authenticated API call to FlexERP endpoint
4. Return inventory data for use in PDF generation
"""

import re
import requests
from typing import Optional, Dict, Any
import config


def parse_barcode(barcode_input: str) -> Optional[str]:
    """
    Extract numeric inventory ID from barcode input.
    
    Args:
        barcode_input (str): Barcode string in format ^12345 or 12345
        
    Returns:
        Optional[str]: Extracted numeric inventory ID, or None if invalid format
        
    Examples:
        >>> parse_barcode("^12345")
        "12345"
        >>> parse_barcode("12345")
        "12345"
        >>> parse_barcode("^ABC123")
        "123"
        
    Note:
        This extracts the first sequence of consecutive digits found in the input.
        For barcodes with multiple numeric sequences (e.g., ABC123DEF456), 
        only the first sequence (123) will be extracted.
    """
    if not barcode_input:
        return None
    
    # Remove leading ^ if present
    cleaned = barcode_input.lstrip('^').strip()
    
    # Extract first sequence of consecutive digits
    match = re.search(r'\d+', cleaned)
    if match:
        return match.group()
    
    return None


def fetch_inventory_data(inventory_id: str) -> Dict[str, Any]:
    """
    Fetch inventory data from FlexERP API for the given inventory ID.
    
    Args:
        inventory_id (str): Numeric inventory ID to look up
        
    Returns:
        Dict[str, Any]: Dictionary containing inventory data with keys:
            Core fields (mapped to PDF template):
            - inventory_id: The inventory ID
            - description: Item description → hoist_desc
            - manufacturer: Manufacturer name → manufacturer
            - model: Model number → model
            - serial_number: Serial number → serial_no
            - power_supply: Power specifications → power_supply
            - rated_capacity: Rated capacity → rated_cap
            
            Optional fields (if available in FlexERP):
            - owner: Customer/owner name → owner
            - location: Location/address → address
            - current_rating: Manufacturer current spec → man_spec_i
            - specifications: Additional specifications (dict)
            
            Status fields:
            - success: Boolean indicating API call success
            - error: Error message if success is False
            - raw_data: Complete API response for reference
            
    Note:
        This function makes a synchronous HTTP request to the FlexERP API.
        Field mappings align with the AVL Letterhead.pdf template fields.
        
    TODO: Complete authentication implementation based on FlexERP API requirements.
          Current implementation uses placeholder authentication.
    """
    if not inventory_id:
        return {
            'success': False,
            'error': 'Invalid inventory ID provided'
        }
    
    # Get API configuration
    base_url = config.get_api_base_url()
    credentials = config.get_flexerp_credentials()
    
    # Validate credentials are configured
    is_valid, error_msg = config.validate_credentials()
    if not is_valid:
        return {
            'success': False,
            'error': error_msg
        }
    
    # TODO: Update endpoint path based on actual FlexERP API documentation
    # This is a placeholder endpoint structure
    endpoint = f"{base_url}inventory/{inventory_id}"
    
    # TODO: Configure authentication headers based on FlexERP API requirements
    # Options may include: Basic Auth, Bearer Token, API Key, etc.
    # Note: Only one authentication method should be used at a time
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    
    # Determine authentication method to use
    # Priority: API Key > Username/Password
    auth = None
    if credentials['api_key']:
        # API Key authentication (adjust header name as needed)
        headers['X-API-Key'] = credentials['api_key']
    elif credentials['username'] and credentials['password']:
        # Will use Basic Auth in the request below
        auth = (credentials['username'], credentials['password'])
    
    try:
        # Make API request with configured timeout and retry logic
        response = requests.get(
            endpoint,
            headers=headers,
            auth=auth,
            timeout=config.API_TIMEOUT
        )
        
        # Check if request was successful
        if response.status_code == 200:
            data = response.json()
            
            # TODO: Map FlexERP API response fields to internal field names
            # Adjust these mappings based on actual API response structure
            return {
                'success': True,
                'inventory_id': inventory_id,
                'description': data.get('description', ''),
                'manufacturer': data.get('manufacturer', ''),
                'model': data.get('model', ''),
                'serial_number': data.get('serial_number', ''),
                'power_supply': data.get('power_supply', ''),
                'rated_capacity': data.get('rated_capacity', ''),
                'owner': data.get('owner', ''),  # Customer/owner information
                'location': data.get('location', ''),  # Location/address
                'current_rating': data.get('current_rating', ''),  # Manufacturer current spec
                'specifications': data.get('specifications', {}),
                'raw_data': data  # Preserve original response for reference
            }
        elif response.status_code == 404:
            return {
                'success': False,
                'error': f'Inventory ID {inventory_id} not found in FlexERP system'
            }
        elif response.status_code == 401:
            return {
                'success': False,
                'error': 'Authentication failed. Please check FlexERP credentials in config.py or .env file'
            }
        else:
            return {
                'success': False,
                'error': f'API request failed with status code {response.status_code}: {response.text}'
            }
            
    except requests.exceptions.Timeout:
        return {
            'success': False,
            'error': f'API request timed out after {config.API_TIMEOUT} seconds'
        }
    except requests.exceptions.ConnectionError:
        return {
            'success': False,
            'error': 'Failed to connect to FlexERP API. Please check network connection and API URL'
        }
    except requests.exceptions.RequestException as e:
        return {
            'success': False,
            'error': f'API request error: {str(e)}'
        }
    except Exception as e:
        return {
            'success': False,
            'error': f'Unexpected error fetching inventory data: {str(e)}'
        }


def fetch_from_barcode(barcode_input: str) -> Dict[str, Any]:
    """
    Complete workflow: parse barcode and fetch inventory data.
    
    This is the main entry point for barcode-based inventory lookup.
    
    Args:
        barcode_input (str): Barcode string (e.g., "^12345" or "12345")
        
    Returns:
        Dict[str, Any]: Inventory data dictionary or error information
        
    Example:
        >>> result = fetch_from_barcode("^12345")
        >>> if result['success']:
        ...     print(f"Found: {result['description']}")
        ... else:
        ...     print(f"Error: {result['error']}")
    """
    inventory_id = parse_barcode(barcode_input)
    
    if not inventory_id:
        return {
            'success': False,
            'error': 'Invalid barcode format. Expected format: ^12345 or 12345'
        }
    
    return fetch_inventory_data(inventory_id)
