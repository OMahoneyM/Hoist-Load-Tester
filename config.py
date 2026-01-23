"""
Configuration module for FlexERP API credentials and settings.

This module manages secure credential storage and retrieval for the FlexERP API integration.
Users should configure their credentials in a .env file or directly update the default values below.
"""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file if it exists
env_path = Path(__file__).parent / '.env'
if env_path.exists():
    load_dotenv(env_path)

# TODO: Update these default values or use environment variables via a .env file
# FlexERP API Configuration
FLEXERP_API_BASE_URL = os.getenv(
    'FLEXERP_API_BASE_URL', 
    'https://flex-solutions-inc.github.io/FlexERP-Public/'
)

# TODO: Enter your FlexERP credentials here or set them as environment variables
# These credentials will be used for API authentication
FLEXERP_USERNAME = os.getenv('FLEXERP_USERNAME', '')
FLEXERP_PASSWORD = os.getenv('FLEXERP_PASSWORD', '')
FLEXERP_API_KEY = os.getenv('FLEXERP_API_KEY', '')

# API Request Settings
API_TIMEOUT = 30  # Timeout for API requests in seconds
API_RETRY_ATTEMPTS = 3  # Number of retry attempts for failed API calls


def get_flexerp_credentials():
    """
    Retrieve FlexERP API credentials from environment or default values.
    
    Returns:
        dict: Dictionary containing API credentials with keys:
            - username: FlexERP username
            - password: FlexERP password
            - api_key: FlexERP API key (if applicable)
    
    Note:
        Users should configure credentials in a .env file or set environment variables.
        For security, never commit credentials directly in code.
    """
    return {
        'username': FLEXERP_USERNAME,
        'password': FLEXERP_PASSWORD,
        'api_key': FLEXERP_API_KEY
    }


def get_api_base_url():
    """
    Get the FlexERP API base URL.
    
    Returns:
        str: The base URL for FlexERP API endpoints
    """
    return FLEXERP_API_BASE_URL


def validate_credentials():
    """
    Validate that required credentials are configured.
    
    Returns:
        tuple: (bool, str) - (is_valid, error_message)
            is_valid: True if credentials are configured, False otherwise
            error_message: Description of missing credentials, empty if valid
    """
    creds = get_flexerp_credentials()
    
    # TODO: Adjust validation based on actual FlexERP authentication requirements
    # This may require username/password, API key, or other authentication methods
    
    # Check for API key authentication
    if creds['api_key']:
        return True, ""
    
    # Check for username/password authentication
    if creds['username'] and creds['password']:
        return True, ""
    
    # Check if username is provided without password
    if creds['username'] and not creds['password']:
        return False, "FLEXERP_USERNAME is set but FLEXERP_PASSWORD is missing in .env file."
    
    # No credentials configured
    return False, "FlexERP credentials not configured. Please set FLEXERP_USERNAME/FLEXERP_PASSWORD or FLEXERP_API_KEY in .env file."
