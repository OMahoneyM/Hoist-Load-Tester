# FlexERP to PDF Field Mapping

This document details how FlexERP inventory data maps to the actual PDF template fields.

## PDF Template Analysis

The AVL Letterhead.pdf template contains **22 form fields**. The FlexERP integration populates these EXISTING fields - no PDF template modification is required.

## Field Mapping Table

| PDF Field Name | Field Type | FlexERP Data Field | Auto-Populated | Notes |
|----------------|------------|-------------------|----------------|-------|
| **Equipment Information** |
| hoist_desc | Text | description | ✓ | Equipment/hoist description |
| manufacturer | Text | manufacturer | ✓ | Manufacturer name |
| model | Text | model | ✓ | Model number |
| serial_no | Text | serial_number | ✓ | Serial number |
| power_supply | Text | power_supply | ✓ | Power supply specifications |
| rated_cap | ComboBox | rated_capacity | ✓ | Rated capacity (matches dropdown: 1/4 ton, 1/2 ton, 1 ton, 2 ton) |
| **Customer Information** |
| owner | Text | owner | ✓ | Customer/owner name (if available in FlexERP) |
| address | Text | location | ✓ | Customer location/address (if available in FlexERP) |
| tester | Text | - | ✗ | Technician name (manual entry) |
| **Test Specifications** |
| load_test_spec | Text | - | ✗ | Load test specification (manual entry) |
| pounds | Text | - | ✗ | Test weight in pounds (manual entry) |
| overload | CheckBox | - | ✗ | Overload protection test result (manual entry) |
| man_spec_i | Text | current_rating | ✓ | Manufacturer's current specification (if available in FlexERP) |
| **Test Results (from Modbus)** |
| actual_i_p1 | Text | - | ✗ | Actual current phase 1 (from Modbus device) |
| actual_i_p2 | Text | - | ✗ | Actual current phase 2 (from Modbus device) |
| actual_i_p3 | Text | - | ✗ | Actual current phase 3 (from Modbus device) |
| **Additional Information** |
| comments | Text | inventory_id | ✓ | Inventory ID prepended to user comments |
| **Date Fields (Auto-filled)** |
| full_date | Text | - | ✗ | Auto-filled with current date |
| date | Text | - | ✗ | Auto-filled with current day |
| month | Text | - | ✗ | Auto-filled with current month |
| year | Text | - | ✗ | Auto-filled with current year |

## FlexERP API Response Structure

The FlexERP API integration expects the following data structure:

```json
{
  "description": "Electric Chain Hoist",
  "manufacturer": "ACME Manufacturing",
  "model": "ECH-2000",
  "serial_number": "SN123456789",
  "power_supply": "480V 3-Phase",
  "rated_capacity": "1 ton",
  "owner": "ABC Industries",
  "location": "123 Main St, City, State",
  "current_rating": "15.5A",
  "specifications": {
    "lift_height": "20 ft",
    "lift_speed": "8 fpm"
  }
}
```

## Field Population Behavior

### Auto-Population Rules

1. **Only Empty Fields**: FlexERP data only populates fields that are currently empty
2. **User Input Priority**: If a user has already entered data, it won't be overwritten
3. **Dropdown Matching**: For rated_capacity, the system attempts to match the exact text in the dropdown
4. **Inventory ID**: Always prepended to the comments field (even if comments exist)

### Manual Entry Required

The following fields always require manual entry:
- Tester name (technician performing the test)
- Load test specification
- Test weight (pounds)
- Overload protection test result
- Actual current readings (obtained from Modbus device during test)

## Integration Workflow

1. **User scans/enters barcode** → System extracts inventory ID
2. **API call to FlexERP** → Fetches inventory data
3. **Auto-population** → Fills empty form fields with FlexERP data
4. **User review/edit** → User can modify any auto-populated fields
5. **Load test** → User runs test, Modbus captures current readings
6. **PDF generation** → All data (FlexERP + manual + Modbus) saved to PDF

## Benefits

- **Reduces data entry errors** - Critical equipment info comes directly from inventory system
- **Maintains existing workflow** - No changes to PDF template or testing procedure
- **Flexible configuration** - Works with or without FlexERP integration
- **Backward compatible** - Existing PDFs and templates work unchanged

## Customization

To add additional field mappings:

1. Add the field to the FlexERP API response in `flexerp_api.py` (line ~135)
2. Add auto-population logic in `main.py` `fetch_flexerp_data()` method (line ~230)
3. Ensure the PDF template has a corresponding form field with the correct name

---

*This mapping is based on the actual AVL Letterhead.pdf template analyzed on 2026-01-29.*
