# RPA SharePoint Connector

A lightweight, production-ready Python connector for SharePoint and OneDrive, designed for RPA bots and automation workflows.

**Perfect for:** UiPath, Automation Anywhere, Blue Prism, custom schedulers, and Python-based automation.

## Features

✨ **Simple Setup** - One-time OAuth authentication, no API keys or secrets  
🔒 **Secure** - Encrypted local token storage, automatic token refresh  
⚡ **Fast** - Optimized uploads up to 100 MB, supports large files  
🔄 **Reliable** - Automatic retry logic, conflict handling, resumable transfers  
🎯 **Flexible** - Works with OneDrive, SharePoint sites, nested folders  
📊 **Bot-Friendly** - JSON output for parsing, non-interactive runtime  

## Quick Start (2 Minutes)

### 1. Install

```bash
pip install -e .
```

### 2. Set Up Authentication (One-Time, Opens Browser)

```bash
python -m rpa_sharepoint_connector configure --profile default --client-id YOUR_APP_ID
```

### 3. Upload a File

```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --folder-url "https://yourcompany.sharepoint.com/Documents/Inbox" \
  --local-path myfile.pdf \
  --remote-path myfile.pdf
```

**Done!** Your file is now in SharePoint.

---

## Installation

### System Requirements
- Python 3.7+
- pip

### Setup

```bash
# Clone repository
git clone https://github.com/yourusername/rpa-sharepoint-connector.git
cd rpa-sharepoint-connector

# Install
pip install -e .

# Development (includes test dependencies)
pip install -e ".[dev]"
```

---

## Azure App Registration (Required)

Before using the connector, register an app in Microsoft Azure:

### Step 1: Create App Registration
1. Go to [Azure Portal](https://portal.azure.com)
2. Search for **"App registrations"**
3. Click **"New registration"**
4. Enter a name (e.g., "SharePoint Connector")
5. Select **"Accounts in any organizational directory"**
6. Click **Register**

### Step 2: Configure Permissions
1. Go to **API permissions**
2. Click **"Add a permission"** → **Microsoft Graph**
3. Select **"Delegated permissions"**
4. Search and add:
   - `offline_access`
   - `User.Read`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
5. Click **Grant admin consent for [Your Organization]**

### Step 3: Configure Redirect URI
1. Go to **Authentication**
2. Click **"Add a platform"** → **Web**
3. Add: `http://localhost/callback`
4. Click **Save**

### Step 4: Copy Your Client ID
1. Go to **Overview**
2. Copy the **Application (client) ID**
3. Use this ID in the setup command

---

## Usage Guide

### Setup (One-Time)

```bash
# Configure with custom client ID
python -m rpa_sharepoint_connector configure \
  --profile default \
  --client-id 73c98460-f8b1-4f61-8938-cd0e0efa462d \
  --force
```

Opens your browser → Sign in → Grant permissions → Done!

Your credentials are **encrypted and saved locally** at `~/.rpa_sharepoint_connector/default.json`.

### Uploading Files

#### Method 1: Simple (Copy-Paste URL)
```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --local-path report.xlsx \
  --remote-path report.xlsx
```

#### Method 2: Flexible (Separate URL and Path)
```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --sharepoint-url "https://company.sharepoint.com" \
  --remote-path "Documents/Inbox/report.xlsx" \
  --local-path report.xlsx
```

#### Handle Conflicts
```bash
python -m rpa_sharepoint_connector run --profile default --op upload \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --local-path file.pdf \
  --remote-path file.pdf \
  --conflict rename  # Options: overwrite, fail_if_exists, rename
```

### Downloading Files

```bash
python -m rpa_sharepoint_connector run --profile default --op download \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --remote-path myfile.pdf \
  --local-path ./myfile.pdf
```

### Listing Files

```bash
python -m rpa_sharepoint_connector run --profile default --op list \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --json  # Machine-readable JSON output
```

Output:
```json
{
  "operation": "list",
  "success": true,
  "folder_path": "Documents/Inbox",
  "count": 3,
  "items": [
    {
      "name": "report.pdf",
      "id": "01BV6CNHTEDH73YZJMLZHK3TQVOT6XRFHR",
      "size": 1024000,
      "is_folder": false
    }
  ]
}
```

### Checking if File Exists

```bash
python -m rpa_sharepoint_connector run --profile default --op exists \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --remote-path myfile.pdf \
  --json
```

### Moving Files

```bash
python -m rpa_sharepoint_connector run --profile default --op move \
  --sharepoint-url "https://company.sharepoint.com" \
  --source-path "Documents/Inbox/myfile.pdf" \
  --target-path "Documents/Processed"
```

### Creating Folders

```bash
python -m rpa_sharepoint_connector run --profile default --op mkdir \
  --sharepoint-url "https://company.sharepoint.com" \
  --folder-path "Documents/NewFolder/SubFolder"
```

### Deleting Files

```bash
python -m rpa_sharepoint_connector run --profile default --op delete \
  --folder-url "https://company.sharepoint.com/Documents/Inbox" \
  --remote-path myfile.pdf
```

---

## Python API

For Python-based bots (RPA Framework, Python scripts):

```python
from rpa_sharepoint_connector import SharePointClient

# Initialize (uses saved credentials)
sp = SharePointClient(profile="default")

# Upload
sp.upload("invoice.pdf", "Invoices/Incoming/invoice.pdf")

# Download
sp.download("Invoices/Incoming/invoice.pdf", "local_invoice.pdf")

# List files
files = sp.list("Invoices/Incoming")
for file in files:
    print(f"{file['name']} - {file['size']} bytes")

# Check existence
exists = sp.exists("Invoices/Incoming/invoice.pdf")
print(f"File exists: {exists}")

# Move
sp.move("Invoices/Incoming/invoice.pdf", "Invoices/Processed")

# Create folder
sp.mkdir("Invoices/Archive/2024")

# Delete
sp.delete("Invoices/Processed/old_file.pdf")

# Close connection
sp.close()
```

---

## File Size Support

| Size | Status | Speed | Timeout |
|------|--------|-------|---------|
| < 1 MB | ✅ Fast | ~0.3 MB/s | 30s |
| 1-10 MB | ✅ Good | ~0.4 MB/s | 1m |
| 10-50 MB | ✅ Supported | ~1-1.6 MB/s | 5m |
| 50-100 MB | ✅ Supported | ~1.5 MB/s | 5m |
| > 100 MB | ⚠️ Consider chunking | - | - |

All uploads include **automatic retry logic** for network failures.

---

## Managing Profiles

### List Profiles
```bash
python -m rpa_sharepoint_connector list
```

### Check Profile Status
```bash
python -m rpa_sharepoint_connector status --profile default
```

### Update Profile
```bash
python -m rpa_sharepoint_connector configure --profile default --force
```

### Disconnect (Delete Profile)
```bash
python -m rpa_sharepoint_connector disconnect --profile default --yes
```

---

## Troubleshooting

### "Profile not found"
**Solution:** Configure first
```bash
python -m rpa_sharepoint_connector configure --profile default --client-id YOUR_APP_ID
```

### "Token expired"
**Solution:** Reconfigure with --force
```bash
python -m rpa_sharepoint_connector configure --profile default --client-id YOUR_APP_ID --force
```

### "Access denied" during setup
**Solution:** Check Azure app permissions
1. Go to Azure Portal → App registrations
2. Check "API permissions" → verify all 4 permissions are granted
3. Click "Grant admin consent"
4. Try setup again

### "File already exists"
**Solution:** Use conflict handling
```bash
--conflict rename  # Creates file (1).pdf if file.pdf exists
--conflict overwrite  # Replaces existing file (default)
--conflict fail_if_exists  # Raises error if exists
```

### "Folder not found"
**Solution:** Ensure folder exists
```bash
# List folders first
python -m rpa_sharepoint_connector run --profile default --op list \
  --sharepoint-url "https://company.sharepoint.com" \
  --folder-path "Documents"

# Create if missing
python -m rpa_sharepoint_connector run --profile default --op mkdir \
  --sharepoint-url "https://company.sharepoint.com" \
  --folder-path "Documents/NewFolder"
```

---

## Real-World Examples

### UiPath RPA Bot
```
Invoke PowerShell Activity:
  Command: python -m rpa_sharepoint_connector run --profile default --op upload --folder-url "https://..." --local-path "C:\temp\file.pdf" --remote-path "file.pdf"
  
Invoke PowerShell Activity:
  Command: python -m rpa_sharepoint_connector run --profile default --op list --folder-url "https://..." --json
  Output: save to ListOutput variable
  
Deserialize JSON Activity:
  Input: ListOutput
  Output: ParsedList
```

### Python Bot (RPA Framework)
```python
from RPA.Robotic_process_automation import RPA
import subprocess
import json

def upload_to_sharepoint(file_path, remote_path):
    result = subprocess.run([
        'python', '-m', 'rpa_sharepoint_connector', 'run',
        '--profile', 'default',
        '--op', 'upload',
        '--folder-url', 'https://company.sharepoint.com/Documents/Bot',
        '--local-path', file_path,
        '--remote-path', remote_path
    ], capture_output=True, text=True)
    
    if result.returncode == 0:
        print("Upload successful!")
    else:
        print(f"Upload failed: {result.stderr}")

def list_sharepoint_files():
    result = subprocess.run([
        'python', '-m', 'rpa_sharepoint_connector', 'run',
        '--profile', 'default',
        '--op', 'list',
        '--folder-url', 'https://company.sharepoint.com/Documents/Bot',
        '--json'
    ], capture_output=True, text=True)
    
    files = json.loads(result.stdout)
    return files['items']
```

### Command-Line Batch Processing
```bash
#!/bin/bash
# Process all PDF files in a folder

SHAREPOINT_URL="https://company.sharepoint.com/Documents/Inbox"
LOCAL_DIR="./files_to_process"

# List files
FILES=$(python -m rpa_sharepoint_connector run --profile default \
  --op list --folder-url "$SHAREPOINT_URL" --json | jq -r '.items[].name')

# Download and process each
for file in $FILES; do
  echo "Processing: $file"
  
  # Download
  python -m rpa_sharepoint_connector run --profile default \
    --op download --folder-url "$SHAREPOINT_URL" \
    --remote-path "$file" --local-path "$LOCAL_DIR/$file"
  
  # Process (your business logic here)
  # ...
  
  # Move to processed folder
  python -m rpa_sharepoint_connector run --profile default \
    --op move --sharepoint-url "https://company.sharepoint.com" \
    --source-path "Documents/Inbox/$file" \
    --target-path "Documents/Processed"
  
  echo "✓ Completed: $file"
done
```

---

## Environment Variables

```bash
# Override default client ID
export MICROSOFT_CLIENT_ID=73c98460-f8b1-4f61-8938-cd0e0efa462d

# Override default tenant
export MICROSOFT_TENANT_ID=organizations

# Override token storage location
export RPA_SHAREPOINT_STORE_DIR=/custom/path/to/tokens
```

---

## Testing

```bash
# Run all tests
python -m pytest -v

# Run specific test
python -m pytest rpa_sharepoint_connector/tests/test_upload_idempotency.py -v

# Run with coverage
python -m pytest --cov=rpa_sharepoint_connector
```

---

## Architecture

- **`auth.py`** - Token refresh and Microsoft auth
- **`browser_auth.py`** - Browser-based OAuth (PKCE) 
- **`graph_client.py`** - Microsoft Graph API wrapper
- **`sdk.py`** - High-level SharePoint client
- **`cli.py`** - Command-line interface
- **`cli_setup.py`** - Setup and configuration
- **`token_store.py`** - Encrypted credential storage

---

## Security

✅ **Credentials encrypted** locally using system keyring  
✅ **No secrets in code** - uses OAuth instead of API keys  
✅ **PKCE flow** - secure browser-based authentication  
✅ **Token auto-refresh** - old tokens automatically refreshed  
✅ **No data logging** - file contents never logged  

---

## Contributing

We welcome contributions! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup
```bash
git clone <your-fork>
cd rpa-sharepoint-connector
pip install -e ".[dev]"
pytest
```

---

## License

MIT License - see LICENSE file for details

---

## Support

- 📖 [Documentation](README.md)
- 🐛 [Report Issues](https://github.com/yourusername/rpa-sharepoint-connector/issues)
- 💬 [Discussions](https://github.com/yourusername/rpa-sharepoint-connector/discussions)

---

## Changelog

### v1.1.0 (Latest)
- ✨ Add `--folder-url` parameter for simplified usage
- ✨ Support both URL formats (simple and flexible)
- 🚀 Increase file upload limit to 100 MB
- ⏱️ Extend timeout to 5 minutes for large files
- 🐛 Fix URL encoding in upload sessions
- 📝 Add comprehensive documentation

### v1.0.0
- Initial release
- One-command setup
- All core operations (upload, download, list, delete, move, mkdir, exists)

---

**Made with ❤️ for RPA teams worldwide**
