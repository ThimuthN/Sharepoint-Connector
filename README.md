# RPA SharePoint Connector

A simple Python tool for automating file operations in SharePoint and OneDrive. Works with UiPath, Blue Prism, automation scripts, and anything that can run Python.

Whether you're building a bot to process invoices, archive documents, or sync files between systems, this tool handles the SharePoint part.

## What it does

- **Upload/download files** - Push files to SharePoint, pull them back down
- **List folders** - See what's in a SharePoint folder with JSON output
- **Move, delete, create folders** - Automate file organization
- **Handle conflicts** - Overwrite, rename, or fail based on what you need
- **Large files** - Works with files up to 100 MB
- **No browser at runtime** - Set it up once, then it runs without opening browser windows

## Why use this

- **Simple auth** - Just one-time browser login, no API keys to manage
- **Works everywhere** - UiPath, Blue Prism, shell scripts, Python code
- **Reliable** - Auto-retries on network failures
- **Secure** - Credentials encrypted locally, never sent to external servers

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

## Set up in Azure (One Time)

You need to register an app in Azure first. Here's how:

1. **Create the app**
   - Go to [portal.azure.com](https://portal.azure.com)
   - Search for "App registrations" and click it
   - Click "New registration"
   - Name it (e.g., "My SharePoint Bot")
   - Pick "Accounts in any organizational directory"
   - Click Register

2. **Give it permissions**
   - Go to "API permissions" in the left menu
   - Click "Add a permission" → "Microsoft Graph" → "Delegated permissions"
   - Search for and add these 4:
     - `offline_access` (so it can refresh tokens)
     - `User.Read` (to get your info)
     - `Files.ReadWrite.All` (to read/write files)
     - `Sites.ReadWrite.All` (to access SharePoint sites)
   - Click "Grant admin consent" at the top

3. **Set the callback URL**
   - Go to "Authentication" in the left menu
   - Click "Add a platform" → pick "Web"
   - Add `http://localhost/callback`
   - Click Save

4. **Copy your Client ID**
   - Go to "Overview"
   - Copy the "Application (client) ID"
   - You'll use this in the next step

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

**Profile not found?**
Haven't set up yet. Run:
```bash
python -m rpa_sharepoint_connector configure --profile default --client-id YOUR_APP_ID
```

**Token expired?**
Your login token expired. Just reconfigure:
```bash
python -m rpa_sharepoint_connector configure --profile default --client-id YOUR_APP_ID --force
```

**"Access denied" during setup?**
The Azure app doesn't have the right permissions:
1. Go to [portal.azure.com](https://portal.azure.com)
2. Find your app in "App registrations"
3. Click "API permissions"
4. Make sure all 4 permissions are there (Files.ReadWrite.All, Sites.ReadWrite.All, etc.)
5. Click "Grant admin consent" at the top
6. Try setup again

**File already exists error?**
Use the `--conflict` flag:
```bash
--conflict overwrite   # Replace it (default)
--conflict rename      # Create file (1).pdf instead
--conflict fail_if_exists  # Raise error instead of uploading
```

**Folder not found?**
The folder might not exist. Check what's there:
```bash
python -m rpa_sharepoint_connector run --profile default --op list \
  --sharepoint-url "https://company.sharepoint.com" \
  --folder-path "Documents"
```

Or create it:
```bash
python -m rpa_sharepoint_connector run --profile default --op mkdir \
  --sharepoint-url "https://company.sharepoint.com" \
  --folder-path "Documents/NewFolder"
```

---

## Examples

### UiPath (Drag & drop PowerShell)

In UiPath Designer, use "Invoke PowerShell":

```
# Upload a file
python -m rpa_sharepoint_connector run --profile default --op upload \
  --folder-url "https://company.sharepoint.com/Documents/Bot" \
  --local-path "C:\temp\invoice.pdf" \
  --remote-path "invoice.pdf"

# List files and save JSON to variable
python -m rpa_sharepoint_connector run --profile default --op list \
  --folder-url "https://company.sharepoint.com/Documents/Bot" --json
```

### Python (Basic)

```python
import subprocess
import json

# Upload a file
subprocess.run([
    'python', '-m', 'rpa_sharepoint_connector', 'run',
    '--profile', 'default',
    '--op', 'upload',
    '--folder-url', 'https://company.sharepoint.com/Documents/Inbox',
    '--local-path', 'report.pdf',
    '--remote-path', 'report.pdf'
])

# List files
result = subprocess.run([
    'python', '-m', 'rpa_sharepoint_connector', 'run',
    '--profile', 'default',
    '--op', 'list',
    '--folder-url', 'https://company.sharepoint.com/Documents/Inbox',
    '--json'
], capture_output=True, text=True)

files = json.loads(result.stdout)
for file in files['items']:
    print(f"Found: {file['name']}")
```

### Bash (Batch processing)

Process all PDFs from an inbox folder:

```bash
#!/bin/bash

INBOX="https://company.sharepoint.com/Documents/Inbox"

# Get list of files
FILES=$(python -m rpa_sharepoint_connector run --profile default \
  --op list --folder-url "$INBOX" --json | jq -r '.items[].name')

# Download and process each
for file in $FILES; do
  echo "Processing $file..."
  
  python -m rpa_sharepoint_connector run --profile default \
    --op download --folder-url "$INBOX" \
    --remote-path "$file" --local-path "$file"
  
  # Do your processing here (OCR, extract data, etc.)
  # process_pdf.py "$file"
  
  # Move to completed folder
  python -m rpa_sharepoint_connector run --profile default \
    --op move --sharepoint-url "https://company.sharepoint.com" \
    --source-path "Documents/Inbox/$file" \
    --target-path "Documents/Processed"
  
  echo "✓ Done: $file"
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
