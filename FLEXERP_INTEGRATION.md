# FlexERP Integration Guide

This document provides detailed information about the FlexERP barcode integration added to the Hoist Load Tester application.

## Overview

The FlexERP integration allows users to scan or manually enter barcodes to automatically fetch and populate hoist inventory data from the FlexERP system. This streamlines the load testing workflow by reducing manual data entry.

**Works with both:**
- Compiled executable (.exe) - for end users
- Python scripts - for developers

## Architecture

### New Components

1. **config.py** - Configuration Module
   - Manages FlexERP API credentials
   - Loads settings from .env file or environment variables
   - Validates credential configuration
   - Provides default values for API settings
   - **Automatically detects if running as executable or script**

2. **flexerp_api.py** - API Client Module
   - Parses barcode input to extract inventory IDs
   - Communicates with FlexERP REST API
   - Handles authentication (supports API key or username/password)
   - Provides comprehensive error handling
   - Returns structured inventory data

3. **test_flexerp_integration.py** - Test Script
   - Validates barcode parsing
   - Tests credential configuration
   - Verifies API connectivity
   - Can run without GUI

### Modified Components

1. **main.py** - Main Application
   - Added barcode input field to GUI
   - Added "Fetch FlexERP Data" button
   - Added status label for fetch operations
   - Implemented auto-population of form fields
   - Extended PDF generation to include FlexERP data
   - Maintains backward compatibility

2. **hoist-load-tester.spec** - PyInstaller Build Configuration
   - Updated to include new modules (config, flexerp_api)
   - Bundles .env.example file with executable
   - Includes required dependencies (requests, dotenv)

## Usage Instructions

### For End Users (Using the Executable)

1. **Locate the .env.example file** in the same directory as `hoist-load-tester-v1.4.0.exe`

2. **Copy and rename** `.env.example` to `.env` (in the same directory as the .exe)

3. **Edit the .env file** with a text editor (Notepad, Notepad++, etc.) and add your FlexERP credentials:
   ```
   # Option 1: API Key (preferred)
   FLEXERP_API_KEY=your_api_key_here
   
   # Option 2: Username/Password
   FLEXERP_USERNAME=your_username
   FLEXERP_PASSWORD=your_password
   ```

4. **Save the file** and close the text editor

5. **Double-click the .exe** to launch the application

**Note:** The application will look for the `.env` file in the same directory as the executable.

### For Developers (Using Python Scripts)

1. Copy the environment template:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and configure your FlexERP credentials:
   ```
   # Option 1: API Key (preferred)
   FLEXERP_API_KEY=your_api_key_here
   
   # Option 2: Username/Password
   FLEXERP_USERNAME=your_username
   FLEXERP_PASSWORD=your_password
   ```

3. (Optional) Customize the API endpoint in `.env`:
   ```
   FLEXERP_API_BASE_URL=https://your-flexerp-instance.com/api/
   ```

### Testing the Integration

Run the test script to verify setup:
```bash
python test_flexerp_integration.py
```

This will show:
- ✓ Barcode parsing validation
- ✓ Credential configuration status
- ✓ API connectivity test (if credentials configured)

### Using in the Application

1. Launch the application:
   - **Executable:** Double-click `hoist-load-tester-v1.4.0.exe`
   - **Python:** Run `python main.py`
   ```bash
   python main.py
   ```

2. In the GUI, locate the "Barcode (FlexERP)" field

3. Enter or scan a barcode (formats: `^12345`, `12345`, or any alphanumeric with digits)

4. Click "Fetch FlexERP Data"

5. Review the status message:
   - Green "✓ Data loaded" = success
   - Red "✗ Failed" = error (check credentials or barcode)

6. Form fields will auto-populate with available data

7. Continue with load test and PDF generation as normal

## Barcode Format

The system accepts flexible barcode formats:
- `^12345` - Standard format with caret prefix
- `12345` - Numeric only
- `^ABC123` - Mixed alphanumeric (extracts: 123)
- Any format containing digits (first sequence is extracted)

## PDF Template Integration

### Actual PDF Template Fields

The AVL Letterhead.pdf template contains these form fields:

**Test Information:**
- owner, tester, address
- date, month, year, full_date

**Equipment Information:**
- hoist_desc, manufacturer, model, serial_no
- power_supply, rated_cap

**Test Results:**
- load_test_spec, pounds, overload
- man_spec_i, actual_i_p1, actual_i_p2, actual_i_p3
- comments

### FlexERP Data Mapping

When you fetch data from FlexERP, it automatically populates these existing PDF fields:

| PDF Field | FlexERP Field | Auto-Populated | Notes |
|-----------|---------------|----------------|-------|
| hoist_desc | description | ✓ | Equipment description |
| manufacturer | manufacturer | ✓ | Manufacturer name |
| model | model | ✓ | Model number |
| serial_no | serial_number | ✓ | Serial number |
| power_supply | power_supply | ✓ | Power specifications |
| rated_cap | rated_capacity | ✓ | Capacity (dropdown match) |
| owner | owner | ✓ | Customer name (if available in FlexERP) |
| address | location | ✓ | Location/address (if available in FlexERP) |
| man_spec_i | current_rating | ✓ | Current rating (if available in FlexERP) |
| comments | inventory_id | ✓ | Inventory ID prepended to comments |

**Important:** The integration uses the EXISTING PDF template fields. You do not need to modify your PDF template to use FlexERP integration. The system maps FlexERP data to the appropriate existing fields.

## Customization

### Changing the API Endpoint

Edit `config.py` or set in `.env`:
```python
FLEXERP_API_BASE_URL = "https://your-custom-endpoint.com/api/"
```

### Adjusting Authentication

The API client in `flexerp_api.py` supports:
- API Key authentication (via `X-API-Key` header)
- HTTP Basic Authentication (username/password)

To customize authentication, edit `flexerp_api.py` in the `fetch_inventory_data()` function:

```python
# Look for TODO comments around line 100-120
# Adjust headers and auth method as needed
```

### Field Mapping

To map different FlexERP API response fields, edit `flexerp_api.py` around line 130-145:

```python
return {
    'success': True,
    'inventory_id': inventory_id,
    'description': data.get('your_field_name', ''),  # <-- Customize here
    'manufacturer': data.get('your_mfg_field', ''),
    # ... etc
}
```

## Troubleshooting

### "Credentials not configured" Error

**Cause:** No .env file or missing credentials

**Solution:**
1. Copy `.env.example` to `.env`
2. Add your FlexERP credentials
3. Restart the application

### "Authentication failed" Error

**Cause:** Invalid credentials or wrong auth method

**Solution:**
1. Verify credentials are correct
2. Check if using correct auth method (API key vs username/password)
3. Review FlexERP API documentation for auth requirements

### "Inventory ID not found" Error

**Cause:** The barcode/ID doesn't exist in FlexERP

**Solution:**
1. Verify the barcode is correct
2. Check if item exists in FlexERP system
3. Verify API endpoint is correct

### "Connection Error"

**Cause:** Network issue or wrong API URL

**Solution:**
1. Check internet connectivity
2. Verify `FLEXERP_API_BASE_URL` is correct
3. Test API URL in browser or with curl

## Security Considerations

1. **Never commit .env files** - Already excluded in `.gitignore`
2. **Use environment variables in production** - Don't hardcode credentials
3. **Rotate API keys regularly** - Follow security best practices
4. **Use HTTPS** - Ensure API endpoint uses secure connection
5. **Limit API key permissions** - Use least-privilege principle

## API Reference

### `flexerp_api.parse_barcode(barcode_input)`
Extracts numeric inventory ID from barcode string.

**Parameters:**
- `barcode_input` (str): Barcode in any format

**Returns:**
- (str or None): Extracted numeric ID or None if invalid

### `flexerp_api.fetch_inventory_data(inventory_id)`
Fetches inventory data from FlexERP API.

**Parameters:**
- `inventory_id` (str): Numeric inventory ID

**Returns:**
- (dict): Inventory data with `success` flag and fields

### `flexerp_api.fetch_from_barcode(barcode_input)`
Complete workflow: parse + fetch in one call.

**Parameters:**
- `barcode_input` (str): Barcode string

**Returns:**
- (dict): Inventory data or error information

### `config.validate_credentials()`
Validates FlexERP credential configuration.

**Returns:**
- (tuple): (is_valid, error_message)

## Future Enhancements

Potential improvements for future development:

1. **Batch Import** - Scan multiple barcodes and generate multiple PDFs
2. **Offline Mode** - Cache inventory data for offline operation
3. **Custom Field Mapping** - GUI for mapping FlexERP fields to PDF fields
4. **Barcode Scanner Integration** - Direct USB barcode scanner support
5. **Audit Trail** - Log all API calls and data fetches
6. **Data Validation** - Validate fetched data before auto-population

## Support

For issues specific to:
- **Barcode parsing**: See `flexerp_api.py`
- **Authentication**: See `config.py` and `flexerp_api.py`
- **GUI integration**: See `main.py`
- **Testing**: Run `test_flexerp_integration.py`

All code includes comprehensive TODO comments where customization may be needed.

---

**Last Updated:** January 2026
**Integration Version:** 1.0
