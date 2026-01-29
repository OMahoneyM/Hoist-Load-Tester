# Hoist Load Tester
This software is used to connect to a networked modbus and record current draw from a 3-phase hoist motor under load. It also allows the user to generate a pdf report with hoist specs and amperage data.

## Features

- **Modbus Integration**: Connect to networked Modbus devices to measure 3-phase current draw
- **FlexERP Barcode Integration**: Scan or manually enter barcodes to fetch inventory data from FlexERP API
- **Auto-Population**: Automatically populate form fields with FlexERP inventory data
- **PDF Report Generation**: Generate professional PDF reports with hoist specifications and test results
- **User-Friendly GUI**: Qt-based interface for easy data entry and testing

## How It's Made:

**Tech used:** Python, PySide6, PyModBus, PyMuPDF, PyInstaller, Requests

## For End Users (Using the Compiled Executable)

### Download and Run

1. Download the latest `hoist-load-tester-v1.4.0.exe` from the releases page
2. Double-click the `.exe` file to launch the application
3. The application will open with the familiar interface

### Configure FlexERP Integration (Optional)

To enable FlexERP barcode integration:

1. Locate the `.env.example` file in the same directory as the `.exe` file
2. Copy `.env.example` and rename it to `.env` (in the same directory as the .exe)
3. Open `.env` with a text editor (Notepad, Notepad++, etc.)
4. Add your FlexERP credentials:
   ```
   FLEXERP_API_BASE_URL=https://flex-solutions-inc.github.io/FlexERP-Public/
   FLEXERP_USERNAME=your_username_here
   FLEXERP_PASSWORD=your_password_here
   # OR
   FLEXERP_API_KEY=your_api_key_here
   ```
5. Save the file and restart the application

**Note:** FlexERP integration is optional. The application works normally without it.

## For Developers

### Setup and Installation

#### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

#### 2. Configure FlexERP Integration (Optional)

The application includes optional FlexERP API integration for barcode-based inventory lookup.

**To configure FlexERP credentials:**

1. Copy the example environment file:
   ```bash
   cp .env.example .env
   ```

2. Edit `.env` and add your FlexERP credentials:
   ```
   FLEXERP_API_BASE_URL=https://flex-solutions-inc.github.io/FlexERP-Public/
   FLEXERP_USERNAME=your_username_here
   FLEXERP_PASSWORD=your_password_here
   # OR
   FLEXERP_API_KEY=your_api_key_here
   ```

3. **Important**: Never commit the `.env` file to version control! It's already in `.gitignore`.

**Alternative: Direct Configuration**

You can also configure credentials directly in `config.py` by editing the default values (not recommended for production).

#### 3. Run the Application

```bash
python main.py
```

#### 4. Building the Executable

To create a standalone executable for distribution:

```bash
# Ensure all dependencies are installed
pip install -r requirements.txt

# Build the executable using PyInstaller
pyinstaller hoist-load-tester.spec

# The compiled executable will be in the dist/ directory
# dist/hoist-load-tester-v1.4.0.exe
```

**Important:** When distributing the executable:
- Include the `.env.example` file in the same directory as the .exe
- Users can copy and rename it to `.env` to configure FlexERP credentials
- The application will automatically look for the `.env` file next to the executable

## Using FlexERP Barcode Integration

1. **Enter Barcode**: In the "Barcode (FlexERP)" field, enter or scan a barcode in the format `^12345` or `12345`
2. **Fetch Data**: Click the "Fetch FlexERP Data" button
3. **Review**: The application will fetch inventory data and auto-populate available form fields
4. **Continue**: Proceed with your load test and PDF generation as normal

**Barcode Format:**
- Accepted formats: `^12345`, `12345`, or any format containing numeric digits
- The numeric portion will be extracted as the inventory ID

**Note:** FlexERP integration is optional. The application works normally without it.

## Testing FlexERP Integration

A test script is provided to verify the FlexERP integration without running the full GUI:

```bash
python test_flexerp_integration.py
```

This will test:
- Barcode parsing with various formats
- Credential validation
- API connectivity (if credentials are configured)

## Configuration Notes

### FlexERP API Endpoints

The FlexERP API integration is configured in `config.py` and `flexerp_api.py`. 

**To update API endpoints or authentication:**

1. Edit `config.py` to change the base URL:
   ```python
   FLEXERP_API_BASE_URL = "https://your-flexerp-instance.com/api/"
   ```

2. Edit `flexerp_api.py` if authentication method needs adjustment:
   - See TODO comments in the `fetch_inventory_data()` function
   - Adjust headers and authentication based on your FlexERP setup

### PDF Template Fields

The PDF template (AVL Letterhead.pdf) contains the following form fields that can be populated:

**Standard Testing Fields:**
- owner, tester, address
- hoist_desc, manufacturer, model, serial_no, power_supply
- rated_cap, load_test_spec, pounds, overload
- man_spec_i, actual_i_p1, actual_i_p2, actual_i_p3
- comments, date, month, year, full_date

**FlexERP Integration - Field Mapping:**

When FlexERP data is fetched, it automatically populates the following PDF fields:

| PDF Field | FlexERP Data Field | Description |
|-----------|-------------------|-------------|
| hoist_desc | description | Hoist/equipment description |
| manufacturer | manufacturer | Manufacturer name |
| model | model | Model number |
| serial_no | serial_number | Serial number |
| power_supply | power_supply | Power supply specifications |
| rated_cap | rated_capacity | Rated capacity (1/4 ton, 1/2 ton, etc.) |
| owner | owner | Customer/owner name (if available) |
| address | location | Customer location/address (if available) |
| man_spec_i | current_rating | Manufacturer current specification (if available) |
| comments | inventory_id | Inventory ID prepended to comments |

**Note:** FlexERP fields are mapped to existing PDF form fields. The integration does not require any modification to your existing PDF template.

## Optimization:


## Lessons Learned:
