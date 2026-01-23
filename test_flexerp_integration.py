#!/usr/bin/env python3
"""
Test script for FlexERP barcode integration.

This script demonstrates the barcode parsing and API integration workflow
without requiring the full GUI application.

Usage:
    python test_flexerp_integration.py
"""

import flexerp_api
import config


def test_barcode_parsing():
    """Test barcode parsing with various formats."""
    print("=" * 60)
    print("Testing Barcode Parsing")
    print("=" * 60)
    
    test_cases = [
        ("^12345", "12345", "Standard format with caret"),
        ("12345", "12345", "Numeric only"),
        ("^ABC123", "123", "Mixed alphanumeric"),
        ("^0000042", "0000042", "Leading zeros"),
        ("", None, "Empty string"),
        ("ABC", None, "No digits"),
    ]
    
    for barcode, expected, description in test_cases:
        result = flexerp_api.parse_barcode(barcode)
        status = "✓" if result == expected else "✗"
        print(f"{status} {description}")
        print(f"  Input: '{barcode}' -> Output: {result} (Expected: {expected})")
    
    print()


def test_credential_validation():
    """Test credential validation."""
    print("=" * 60)
    print("Testing Credential Validation")
    print("=" * 60)
    
    is_valid, message = config.validate_credentials()
    
    if is_valid:
        print("✓ Credentials are configured and valid")
        print(f"  API Base URL: {config.get_api_base_url()}")
    else:
        print("✗ Credentials are not configured")
        print(f"  Error: {message}")
        print()
        print("To configure credentials:")
        print("  1. Copy .env.example to .env")
        print("  2. Edit .env and add your FlexERP credentials")
        print("  3. Run this test again")
    
    print()


def test_api_call():
    """Test API call with a sample barcode (will fail without valid credentials)."""
    print("=" * 60)
    print("Testing FlexERP API Call")
    print("=" * 60)
    
    # Check if credentials are configured
    is_valid, message = config.validate_credentials()
    if not is_valid:
        print("⚠ Skipping API test - credentials not configured")
        print(f"  {message}")
        print()
        return
    
    # Test with a sample barcode
    test_barcode = "^12345"
    print(f"Testing API call with barcode: {test_barcode}")
    
    result = flexerp_api.fetch_from_barcode(test_barcode)
    
    if result.get('success'):
        print("✓ API call successful!")
        print(f"  Inventory ID: {result.get('inventory_id')}")
        print(f"  Description: {result.get('description')}")
        print(f"  Manufacturer: {result.get('manufacturer')}")
        print(f"  Model: {result.get('model')}")
    else:
        print("✗ API call failed")
        print(f"  Error: {result.get('error')}")
    
    print()


def main():
    """Run all tests."""
    print()
    print("FlexERP Integration Test Suite")
    print("=" * 60)
    print()
    
    test_barcode_parsing()
    test_credential_validation()
    test_api_call()
    
    print("=" * 60)
    print("Test suite complete!")
    print()


if __name__ == "__main__":
    main()
