# Microsoft Graph Utilities

This directory contains utility scripts and tools for Microsoft Graph operations.

## Scripts

### Python Utilities
- `msgraph_upn_lookup_cert.py` - Python script for UPN lookup using certificate authentication

## Usage Examples

### Python UPN Lookup
```bash
# Run Python UPN lookup with certificate authentication
python3 msgraph_upn_lookup_cert.py --email "user@domain.com"
```

## Requirements

### Python Dependencies
For the Python utilities, install required packages:
```bash
pip install msal requests
```

### Certificates
Python scripts use the same certificate authentication as PowerShell scripts, configured in the main `Config.ps1` file.

## Integration

These utilities complement the PowerShell scripts and provide alternative implementation approaches for:
- Cross-platform compatibility (Python runs on Linux/Mac/Windows)
- Integration with existing Python workflows
- Alternative authentication methods
- Custom API operations not available in PowerShell modules