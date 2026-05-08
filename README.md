# Microsoft SharePoint Connector MVP

Minimal production-shaped connector for browsing and downloading files from Microsoft SharePoint using Microsoft Graph API.

**Stack:** FastAPI (Python) backend, React frontend, SQLite storage (MVP), OAuth 2.0.

## Features

- **OAuth 2.0 Authentication** - Delegate user auth via Microsoft identity platform
- **SharePoint Site Browsing** - Search and select SharePoint sites
- **Document Library Navigation** - Browse drives and folders
- **File Download** - Stream file downloads from SharePoint
- **Token Management** - Automatic token refresh, encryption at rest
- **Error Handling** - Clean error messages, no secret leakage
- **Minimal** - Single user, delegated auth only, no advanced features

## MVP Limitations

This MVP intentionally excludes:
- **App-only auth** (client credentials) - delegated user auth only
- **Multi-tenancy hardening** - uses OAuth common/tenant config, not production-grade
- **Webhook/delta sync** - no event subscriptions
- **Teams-specific browsing** - SharePoint sites/drives only
- **Permission mirroring** - no ACL sync
- **Background workers** - all synchronous
- **Advanced caching** - minimal

**Not for production.** Use this as a reference for building your own connector.

## Quick Start

### Prerequisites

- Python 3.9+
- Node.js 18+
- Azure App Registration with Microsoft Graph delegated permissions

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → App registrations → New registration
2. Name: "SharePoint Connector MVP"
3. Supported account types: "Accounts in this organizational directory only"
4. Redirect URI: `http://localhost:8000/api/integrations/microsoft/callback`
5. Save **Application (client) ID** and **Directory (tenant) ID**
6. Certificates & secrets → New client secret → save **Value**
7. API permissions → Add permission → Microsoft Graph → Delegated permissions:
   - `User.Read`
   - `Sites.Read.All`
   - `Files.Read.All`
   - `offline_access`
8. Admin consent → Grant admin consent (if available)

### 2. Backend Setup

```bash
cd backend

# Create virtual environment
python -m venv venv
source venv/Scripts/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Generate encryption key
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"

# Create .env
cp .env.example .env
# Edit .env with your values:
# MICROSOFT_CLIENT_ID=<your-client-id>
# MICROSOFT_CLIENT_SECRET=<your-client-secret>
# MICROSOFT_TENANT_ID=<your-tenant-id>
# TOKEN_ENCRYPTION_KEY=<generated-key>

# Run tests
pytest -v

# Start server
python -m uvicorn app:app --reload
```

Backend runs at `http://localhost:8000`

### 3. Frontend Setup

```bash
cd frontend

# Install dependencies
npm install

# Run tests
npm test

# Start dev server
npm run dev
```

Frontend runs at `http://localhost:3000`

### 4. Test OAuth Flow

1. Open `http://localhost:3000`
2. Click "Connect Microsoft Account"
3. Sign in with your Microsoft account
4. Grant consent to required permissions
5. Select a SharePoint site → document library → folder → file
6. Download a file to verify end-to-end flow

## API Endpoints

### OAuth Flow
- `GET /api/integrations/microsoft/connect` - Redirect to Microsoft OAuth authorization
- `GET /api/integrations/microsoft/callback?code=...&state=...` - Handle OAuth callback

### Connection
- `GET /api/integrations/microsoft/status` - Check connection status

### SharePoint Browsing
- `GET /api/integrations/microsoft/sites?search=query` - List sites (optional search)
- `GET /api/integrations/microsoft/sites/{site_id}/drives` - List document libraries
- `GET /api/integrations/microsoft/drives/{drive_id}/items?item_id=parent` - List folder contents
- `GET /api/integrations/microsoft/drives/{drive_id}/items/{item_id}/download` - Download file

## Configuration

### Backend Environment Variables

```
# Microsoft OAuth
MICROSOFT_CLIENT_ID=<Azure app registration client ID>
MICROSOFT_CLIENT_SECRET=<Azure app registration client secret>
MICROSOFT_TENANT_ID=<Azure tenant ID or "common">
MICROSOFT_REDIRECT_URI=http://localhost:8000/api/integrations/microsoft/callback

# Encryption
TOKEN_ENCRYPTION_KEY=<base64-encoded Fernet key>

# App
APP_BASE_URL=http://localhost:3000
ENVIRONMENT=development
LOG_LEVEL=info
```

### Frontend Environment Variables

```
REACT_APP_API_URL=http://localhost:8000/api
```

## Testing

### Backend Tests

```bash
cd backend
pytest -v                    # Run all tests
pytest tests/test_oauth.py   # Run specific test file
pytest -k test_name          # Run by test name
```

Tests cover:
- OAuth URL generation with correct scopes and parameters
- State validation on callback
- Token encryption/decryption
- Token refresh (expired, near-expiry, invalid_grant)
- Graph API client request/error handling
- API endpoint behavior

**All tests mock Microsoft Graph.** No real credentials required.

### Frontend Tests

```bash
cd frontend
npm test                    # Run all tests
npm test -- --ui            # Run with UI
```

Tests cover:
- Component rendering (connected/disconnected states)
- Site/drive/item listing
- File download calls
- Navigation and breadcrumb
- Error handling

**All tests mock API calls.** No backend required.

## Architecture

### Backend

```
backend/
├── app.py              # FastAPI entry point
├── config.py           # Config validation
├── encryption.py       # Token encryption/decryption
├── models.py           # Data models (MicrosoftConnection)
├── repositories.py     # In-memory storage
├── oauth.py            # OAuth flow manager
├── token_provider.py   # Token refresh provider
├── graph_client.py     # Microsoft Graph API wrapper
├── routes.py           # API route handlers
├── requirements.txt    # Python dependencies
└── tests/              # Test suite
    ├── test_oauth.py
    ├── test_token_provider.py
    ├── test_graph_client.py
    ├── test_encryption.py
    ├── test_repositories.py
    ├── test_config.py
    └── test_routes.py
```

### Frontend

```
frontend/
├── index.html                    # Entry point
├── src/
│   ├── main.jsx                  # React entry
│   ├── App.jsx                   # Root component
│   ├── App.css                   # Styles
│   ├── MicrosoftConnect.jsx       # Connect button & status
│   ├── FileBrowser.jsx            # Site/drive/file browser
│   └── api.js                    # API client
├── tests/                        # Test suite
│   ├── api.test.js
│   ├── MicrosoftConnect.test.jsx
│   └── FileBrowser.test.jsx
├── package.json
├── vite.config.js
└── vitest.config.js
```

## Error Handling

Handled error scenarios:
- **Missing configuration** → Clear error message, fail fast
- **Invalid OAuth state** → 400 Bad Request
- **Expired/invalid token** → Token provider refreshes or marks disconnected
- **Refresh token invalid** → Connection marked disconnected, requires re-auth
- **Graph API 401/403/404/429** → Descriptive error messages
- **File download denied** → Permission error
- **Network failures** → Logged, user-friendly errors

## Security Notes

⚠️ **MVP, not production-ready:**

1. **Token Storage** - Encrypted with Fernet (symmetric encryption). In production, use key vault.
2. **State Validation** - Stored server-side in memory. In production, use Redis/database.
3. **Session** - Hardcoded demo user. Add real user/session management.
4. **CORS** - Configured for localhost. Restrict in production.
5. **HTTPS** - Not enforced in MVP. Required in production.
6. **Rate Limiting** - Not implemented. Add per-user/IP limits.
7. **Audit Logging** - Basic logging only. Add detailed audit trail.
8. **Refresh Token Rotation** - Not implemented. Add rotation policy.

## Deployment

For production deployment:
1. Use environment-based secrets (Azure KeyVault, GitHub Secrets, etc.)
2. Run tests in CI/CD
3. Use proper database (PostgreSQL, not SQLite)
4. Add request validation and rate limiting
5. Implement audit logging
6. Enable HTTPS/TLS
7. Add health checks and monitoring
8. Separate frontend/backend domains
9. Add CSRF protection
10. Implement refresh token rotation

## Troubleshooting

### "Invalid or expired state"
- Check browser has cookies enabled
- Ensure OAuth callback URL matches Azure registration

### "Missing required configuration"
- Verify all environment variables in `.env`
- Run: `python -c "from config import get_config; get_config().validate()"`

### Token refresh fails with "invalid_grant"
- Refresh token expired (user must reconnect)
- Client credentials incorrect

### "Forbidden" on sites/files listing
- User lacks consent for required permissions
- Admin must grant tenant-wide consent in Azure

### Frontend not communicating with backend
- Verify backend running on port 8000
- Check CORS allowed origins match `APP_BASE_URL`
- Check browser DevTools Network tab for 401/403

## Files Changed Summary

### Backend (14 files, ~1,200 LOC)
- Core: `config.py`, `models.py`, `encryption.py`, `repositories.py`
- OAuth: `oauth.py`
- Token: `token_provider.py`
- Graph: `graph_client.py`
- Routes: `routes.py`, `app.py`
- Config: `requirements.txt`
- Tests: 7 test files in `tests/`

### Frontend (11 files, ~900 LOC)
- Core: `App.jsx`, `App.css`, `api.js`
- Components: `MicrosoftConnect.jsx`, `FileBrowser.jsx`
- Config: `package.json`, `vite.config.js`, `vitest.config.js`
- Tests: 3 test files in `tests/`
- Entry: `index.html`, `main.jsx`

### Total: **1,200+ LOC**, **zero external connectors or frameworks**

## References

- [Microsoft Identity Platform v2 Endpoints](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-protocols)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/overview)
- [FastAPI Documentation](https://fastapi.tiangolo.com/)
- [React Documentation](https://react.dev/)

## License

MIT
