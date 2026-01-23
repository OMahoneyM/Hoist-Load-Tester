# Building the Executable

This document describes how to build the Hoist Load Tester executable using PyInstaller.

## Prerequisites

- Python 3.x (developed with Anaconda)
- All dependencies installed (`pip install -r requirements.txt`)
- PyInstaller installed
- Windows development machine (for .exe compilation)

## Build Steps

1. **Install Dependencies**
   
   Ensure all required packages are installed:
   ```bash
   pip install -r requirements.txt
   ```

2. **Verify the Spec File**
   
   The `hoist-load-tester.spec` file is pre-configured with:
   - All necessary modules (config, flexerp_api, requests, dotenv)
   - The .env.example file for distribution
   - Application icon
   
   Current version: v1.4.0 (includes FlexERP integration)

3. **Build the Executable**
   
   Run PyInstaller with the spec file:
   ```bash
   pyinstaller hoist-load-tester.spec
   ```

4. **Locate the Output**
   
   The compiled executable will be created in:
   ```
   dist/hoist-load-tester-v1.4.0.exe
   ```

5. **Prepare for Distribution**
   
   When distributing the application, include:
   - `hoist-load-tester-v1.4.0.exe` (the main executable)
   - `.env.example` (should be in the same directory as the .exe)
   
   The .env.example file is automatically bundled by PyInstaller and will be
   extracted to the same directory as the executable when it runs.

## Testing the Executable

1. **Test Basic Functionality**
   
   Double-click the .exe to ensure it launches without errors.

2. **Test FlexERP Integration**
   
   - Copy `.env.example` to `.env` in the same directory as the executable
   - Add test credentials to `.env`
   - Launch the executable
   - Try the "Fetch FlexERP Data" feature

## Troubleshooting Build Issues

### Missing Modules
If PyInstaller reports missing modules, add them to the `hiddenimports` list in the spec file.

### Missing Files
If data files are not included, add them to the `datas` list in the spec file using the format:
```python
('source/path/to/file', 'destination/in/exe')
```

### Build Fails
- Ensure all dependencies are installed
- Try cleaning the build cache: `pyinstaller --clean hoist-load-tester.spec`
- Check the build output for specific error messages

## Version History

- **v1.4.0** - Added FlexERP barcode integration
- **v1.3.0** - Previous version (baseline)

## Notes

- The application uses PyInstaller's "onefile" mode, creating a single executable
- The .env file is NOT bundled (for security) - only .env.example is included
- Users must create their own .env file from the example to configure credentials
- The config.py module automatically detects when running as an executable and looks for .env in the correct location
