# RPA SharePoint Connector

Simple Python package for RPA bots to interact with Microsoft SharePoint. No browser login required in bot code. Configure once, use everywhere.

## Architecture

Three modes:
- **Configure**: Interactive browser-based setup (one time)
- **Runtime**: No UI, no auth dialogs. Bot imports and uses simple SDK
- **Repair**: Reconnect, test, disconnect commands

## Quick Start

### 1. Install

```bash
# From connector directory
pip install -e rpa_sharepoint_connector
```

### 2. Configure Profile (One Time)

```bash
python -m rpa_sharepoint_connector configure --profile client_a
```

Opens browser UI where you:
- Enter Microsoft Client ID and Secret
- Authenticate with Microsoft
- Select SharePoint site and document library
- Choose default folder
- Test upload/download
- Save profile

### 3. Run Bot (No UI Required)

```python
from rpa_sharepoint_connector import SharePointClient

# No login prompt - uses encrypted stored token
sp = SharePointClient(profile="client_a")

# Upload
sp.upload("invoice.pdf", "Invoices/Incoming/invoice.pdf")

# Download
sp.download("Invoices/Approved/file.pdf", "local.pdf")

# List folder
files = sp.list("Invoices/Incoming")

# Delete
sp.delete("Invoices/Temp/old.pdf")

# Check existence
if sp.exists("Invoices/2024/January.pdf"):
    sp.download("Invoices/2024/January.pdf", "january.pdf")

# Create folder
sp.mkdir("Invoices/2024/Q1")

# Move/rename
sp.move("Invoices/Temp/file.pdf", "Invoices/Archive")
```

## Azure App Registration

Required one-time setup for your organization:

1. Go to https://portal.azure.com → App registrations
2. Create new app registration
3. Add Redirect URI: `http://localhost:8001/callback`
4. Go to Certificates & secrets → New client secret (save the value)
5. Go to API permissions:
   - Add `User.Read` (Microsoft Graph)
   - Add `Sites.Read.All` (Microsoft Graph)
   - Add `Files.Read.All` (Microsoft Graph)
   - Add `Files.ReadWrite.All` (Microsoft Graph - for uploads)
   - Grant admin consent

6. Copy Client ID and Client Secret

## CLI Commands

### Configure Profile
```bash
python -m rpa_sharepoint_connector configure --profile client_a
```

### Check Status
```bash
python -m rpa_sharepoint_connector status --profile client_a
```

Output:
```
Profile: client_a
============================================================
User: john@company.com
Site: Company - Invoices
Library: Shared Documents
Folder: Invoices/2024
Token expires in: 1247 minutes
```

### Test Upload
```bash
python -m rpa_sharepoint_connector test-upload ./sample.pdf --profile client_a
```

### List Profiles
```bash
python -m rpa_sharepoint_connector list
```

### Disconnect Profile
```bash
python -m rpa_sharepoint_connector disconnect --profile client_a
```

## SDK Methods

### Upload
```python
sp.upload(local_path: str, remote_path: str) -> str
```
Upload file. Returns item ID.

### Download
```python
sp.download(remote_path: str, local_path: str) -> None
```
Download file. Supports both path (`Folder/file.pdf`) and item ID.

### Delete
```python
sp.delete(remote_path: str) -> None
```
Delete file or folder.

### Exists
```python
sp.exists(remote_path: str) -> bool
```
Check if item exists.

### List
```python
sp.list(folder_path: str = "") -> List[Dict]
```
List files and folders. Returns list of `{name, id, size, is_folder}`.

### Create Folder
```python
sp.mkdir(folder_path: str) -> str
```
Create folder(s) recursively. Returns item ID of last folder.

### Move/Rename
```python
sp.move(source_path: str, target_path: str, new_name: str = None) -> None
```
Move file or folder. Optionally rename.

## Token Management

- Tokens stored encrypted in `~/.rpa_sharepoint_connector/`
- Client ID and Secret **not** stored
- Refresh token persisted - auto-refreshes when needed
- If refresh fails, bot raises error (not silent failure)

## Security

- Tokens encrypted with Fernet (symmetric encryption)
- Profile key file: `~/.rpa_sharepoint_connector/.key` (chmod 600)
- Profile files encrypted at rest
- No credentials in bot code
- No credentials in environment variables (unless you add them)
- Client secret only used during config, never stored

## File Paths

Paths support both formats:
- **Folder path**: `"Invoices/Incoming/2024"`
- **Item ID**: `"01ABCD123XYZ"`

SDK converts paths to IDs automatically.

## Limitations (MVP)

- Delegated user auth only (not app-only/client-credentials)
- Single user per profile (one login → one account)
- No multi-tenant production hardening (use common endpoint)
- No webhook support
- No delta sync or change tracking
- No Teams integration
- No advanced permission mirroring
- No background jobs
- No bulk operations (upload 1000 files = 1000 API calls)

Future phases (not included):
- Central connection manager for multiple machines
- Connector framework for Google Drive, SFTP, Outlook
- Webhook sync
- Bulk operations

## Troubleshooting

### Token Expired
```
Error: Token refresh failed - Refresh token is invalid or expired
```
Solution:
```bash
python -m rpa_sharepoint_connector configure --profile client_a
```

### Permission Denied
Ensure app has these permissions (in Azure portal):
- Files.ReadWrite.All (or Files.Read.All for read-only)
- Sites.Read.All
- User.Read

Admin consent required.

### Path Not Found
```python
try:
    sp.upload("file.pdf", "NonExistent/Folder/file.pdf")
except ValueError as e:
    print(f"Folder doesn't exist: {e}")
```

Solution: Create folder first:
```python
sp.mkdir("NonExistent/Folder")
sp.upload("file.pdf", "NonExistent/Folder/file.pdf")
```

### Rate Limited
Microsoft Graph has rate limits. If you get:
```
Error: Rate limited. Try again later.
```

Add retry logic:
```python
import time

for attempt in range(3):
    try:
        sp.upload(...)
        break
    except ValueError as e:
        if "Rate limited" in str(e):
            time.sleep(60)
        else:
            raise
```

## Example Bot

See `example_bot.py` for complete working example.

Run:
```bash
python example_bot.py
```

## Module Architecture

```
rpa_sharepoint_connector/
  __init__.py         # Exports SharePointClient
  __main__.py         # CLI entry point
  sdk.py              # Main bot interface
  auth.py             # Microsoft OAuth
  token_store.py      # Encrypted storage
  graph_client.py     # Graph API wrapper
  profiles.py         # Config management
  config_ui.py        # Browser config tool
  cli.py              # Command-line interface
```

## Development

Run tests:
```bash
pytest rpa_sharepoint_connector/tests -v
```

## Support

Issues with:
- Microsoft Graph API → Check Azure app permissions and admin consent
- Encrypted token errors → Delete profile and reconfigure
- Path not found → List folder to find correct IDs
- Rate limits → Add retry logic with backoff

## License

This is example code for educational purposes.
