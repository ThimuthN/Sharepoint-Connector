# RPA SharePoint Connector

Single setup flow for bots:
- One-time interactive browser login during `configure`
- Encrypted token profile stored locally
- Runtime bot operations run silently (no UI)

## Install

```bash
pip install -e rpa_sharepoint_connector
```

## Recommended Flow

1. One-time setup:

```bash
python -m rpa_sharepoint_connector configure --profile client_a
```

2. Bot runtime:

```python
from rpa_sharepoint_connector import SharePointClient

sp = SharePointClient(profile="client_a")
sp.upload("invoice.pdf", "Invoices/Incoming/invoice.pdf")
```

3. Reconnect if profile is broken or needs replacement:

```bash
python -m rpa_sharepoint_connector configure --profile client_a --force
```

## CLI Commands

```bash
python -m rpa_sharepoint_connector configure --profile client_a
python -m rpa_sharepoint_connector status --profile client_a
python -m rpa_sharepoint_connector test-upload ./sample.pdf --profile client_a
python -m rpa_sharepoint_connector disconnect --profile client_a
```

Optional setup override:

```bash
python -m rpa_sharepoint_connector configure --profile client_a --redirect-uri http://localhost:8765/callback
python -m rpa_sharepoint_connector configure --profile client_a --client-id <app_id> --tenant-id organizations
```

## Authentication Model

- OAuth 2.0 Authorization Code Flow with PKCE
- Local callback server on `http://localhost/callback` (ephemeral localhost port at runtime)
- Strict state validation on callback
- Public-client only (no client secret)
- Refresh token used for silent runtime token renewal
- `client_id` and `tenant_id` used during setup are saved in the encrypted profile and reused at runtime

## Azure App Requirements

Use one public-client app registration:

- Redirect URI: `http://localhost/callback`
- Platform: public client / desktop-compatible localhost redirect
- No client secret
- Delegated permissions:
  - `offline_access`
  - `User.Read`
  - `Files.ReadWrite.All`
  - `Sites.ReadWrite.All`

## Runtime Guarantees

`SharePointClient(profile="client_a")` only:
1. Loads saved profile
2. Refreshes token if needed
3. Calls Graph API

It does not:
- open browser
- trigger OAuth setup
- ask for interactive login
- auto-run configure
