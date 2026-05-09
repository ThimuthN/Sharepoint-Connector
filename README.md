# RPA SharePoint Connector

Lean Python connector for bot workflows:
- One-time browser OAuth setup (Authorization Code + PKCE)
- Encrypted local token/profile storage
- Silent runtime token refresh
- Upload/download/list/delete/move operations via Microsoft Graph

## Install

```bash
pip install -e .
```

## Required Azure App Settings

- App type: public client (no client secret)
- Redirect URI: `http://localhost/callback`
- Delegated permissions:
  - `offline_access`
  - `User.Read`
  - `Files.ReadWrite.All`
  - `Sites.ReadWrite.All`

## CLI

```bash
python -m rpa_sharepoint_connector setup --profile client_a --my-drive
python -m rpa_sharepoint_connector configure --profile client_a
python -m rpa_sharepoint_connector status --profile client_a
python -m rpa_sharepoint_connector set-target --profile client_a --sharepoint-url "<folder-url>"
python -m rpa_sharepoint_connector test-upload ./sample.txt --profile client_a
python -m rpa_sharepoint_connector disconnect --profile client_a
```

Useful overrides:

```bash
python -m rpa_sharepoint_connector setup --profile client_a --sharepoint-url "<folder-url>" --tenant-id organizations
python -m rpa_sharepoint_connector configure --profile client_a --force --client-id <app-id> --tenant-id organizations
python -m rpa_sharepoint_connector set-target --profile client_a --my-drive --folder ConnectorSmoke
```

## Runtime Example

```python
from rpa_sharepoint_connector import SharePointClient

sp = SharePointClient(profile="client_a")
sp.upload("invoice.pdf", "Invoices/Incoming/invoice.pdf")
sp.download("Invoices/Incoming/invoice.pdf", "invoice_local.pdf")
sp.move("Invoices/Incoming/invoice.pdf", "Invoices/Processed")
sp.delete("Invoices/Processed/invoice.pdf")
items = sp.list("Invoices/Incoming")
```

## Testing

```bash
python -m pytest -q
```

## Notes

- Runtime never opens browser.
- If profile is missing or expired, run `configure` again with `--force`.
