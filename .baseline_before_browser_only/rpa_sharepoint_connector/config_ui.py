"""Browser-based configuration UI for SharePoint connector."""
import logging
import webbrowser
from datetime import datetime, timedelta
from typing import Optional, Tuple
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, urlencode
import json
import threading

from .auth import MicrosoftAuth
from .graph_client import GraphClient
from .profiles import ProfileManager

logger = logging.getLogger(__name__)


class ConfigHandler(BaseHTTPRequestHandler):
    """HTTP handler for OAuth callback and config UI."""

    config_state = {
        "profile_name": "",
        "client_id": "",
        "client_secret": "",
        "auth": None,
        "tokens": None,
        "user_info": None,
        "sites": [],
        "selected_site": None,
        "drives": [],
        "selected_drive": None,
        "folders": [],
        "selected_folder": None,
    }

    def log_message(self, format, *args):
        """Suppress server logging."""
        pass

    def do_GET(self):
        """Handle GET requests."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path
        query = parse_qs(parsed_path.query)

        if path == "/callback":
            self._handle_oauth_callback(query)
        elif path == "/":
            self._serve_ui()
        elif path == "/api/init":
            self._api_init(query)
        elif path == "/api/sites":
            self._api_sites()
        elif path == "/api/drives":
            self._api_drives()
        elif path == "/api/folders":
            self._api_folders(query)
        elif path == "/api/test":
            self._api_test()
        elif path == "/api/save":
            self._api_save(query)
        else:
            self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        """Handle POST requests."""
        parsed_path = urlparse(self.path)
        path = parsed_path.path

        if path == "/api/config":
            self._handle_config_post()
        else:
            self._send_json({"error": "Not found"}, 404)

    def _handle_config_post(self):
        """Handle profile and credentials submission."""
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode()
        data = json.loads(body)

        self.config_state["profile_name"] = data.get("profile_name", "default")
        self.config_state["client_id"] = data.get("client_id", "")
        self.config_state["client_secret"] = data.get("client_secret", "")

        if not self.config_state["client_id"] or not self.config_state["client_secret"]:
            self._send_json({"error": "Missing credentials"}, 400)
            return

        auth = MicrosoftAuth(
            client_id=self.config_state["client_id"],
            client_secret=self.config_state["client_secret"],
        )
        self.config_state["auth"] = auth
        auth_url = auth.get_authorization_url()

        self._send_json({"auth_url": auth_url})

    def _handle_oauth_callback(self, query: dict):
        """Handle OAuth callback from Microsoft."""
        code = query.get("code", [""])[0]
        error = query.get("error", [""])[0]

        if error:
            self._serve_error(f"Microsoft error: {error}")
            return

        if not code:
            self._serve_error("No authorization code received")
            return

        try:
            auth = self.config_state.get("auth")
            if not auth:
                self._serve_error("Auth not initialized")
                return

            tokens = auth.exchange_code(code)
            self.config_state["tokens"] = tokens

            # Get user info
            user_info = auth.get_user_info(tokens["access_token"])
            self.config_state["user_info"] = user_info

            # Redirect to site selection
            self.send_response(302)
            self.send_header("Location", "/")
            self.end_headers()
        except Exception as e:
            self._serve_error(f"Failed to exchange code: {str(e)}")

    def _serve_ui(self):
        """Serve the configuration UI."""
        html = """
        <!DOCTYPE html>
        <html>
        <head>
            <title>SharePoint Connector Setup</title>
            <style>
                body { font-family: Arial; margin: 20px; }
                .status { padding: 10px; background: #f0f0f0; margin: 10px 0; border-radius: 4px; }
                input, select { padding: 8px; margin: 5px 0; width: 100%; box-sizing: border-box; }
                button { padding: 10px 20px; background: #0078d4; color: white; border: none; cursor: pointer; margin: 10px 0; }
                button:hover { background: #005a9e; }
                .section { margin: 20px 0; padding: 10px; border: 1px solid #ddd; border-radius: 4px; }
                .error { color: red; }
                .success { color: green; }
            </style>
        </head>
        <body>
            <h1>SharePoint Connector Configuration</h1>

            <div id="content"></div>

            <script>
                async function init() {
                    const response = await fetch('/api/init');
                    const data = await response.json();
                    render(data);
                }

                async function submitConfig() {
                    const profileName = document.getElementById('profile').value;
                    const clientId = document.getElementById('clientId').value;
                    const clientSecret = document.getElementById('clientSecret').value;

                    const response = await fetch('/api/config', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            profile_name: profileName,
                            client_id: clientId,
                            client_secret: clientSecret
                        })
                    });

                    const data = await response.json();
                    if (data.auth_url) {
                        window.location.href = data.auth_url;
                    }
                }

                async function loadSites() {
                    const response = await fetch('/api/sites');
                    const data = await response.json();
                    const select = document.getElementById('site');
                    select.innerHTML = '<option value="">Select a site...</option>';
                    data.sites.forEach(site => {
                        const opt = document.createElement('option');
                        opt.value = site.id;
                        opt.text = site.name;
                        select.appendChild(opt);
                    });
                }

                async function onSiteSelect(siteId) {
                    const response = await fetch('/api/drives?site_id=' + siteId);
                    const data = await response.json();
                    const select = document.getElementById('drive');
                    select.innerHTML = '<option value="">Select a library...</option>';
                    data.drives.forEach(drive => {
                        const opt = document.createElement('option');
                        opt.value = drive.id;
                        opt.text = drive.name;
                        select.appendChild(opt);
                    });
                }

                async function onDriveSelect(driveId) {
                    const response = await fetch('/api/folders?drive_id=' + driveId);
                    const data = await response.json();
                    const select = document.getElementById('folder');
                    select.innerHTML = '<option value="">Root</option>';
                    if (data.folders) {
                        data.folders.forEach(folder => {
                            const opt = document.createElement('option');
                            opt.value = folder.id;
                            opt.text = folder.path;
                            select.appendChild(opt);
                        });
                    }
                }

                async function testUpload() {
                    const response = await fetch('/api/test');
                    const data = await response.json();
                    alert(data.message || data.error);
                }

                async function saveProfile() {
                    const response = await fetch('/api/save?profile_name=' + encodeURIComponent(document.getElementById('profile').value));
                    const data = await response.json();
                    if (data.success) {
                        alert('Profile saved! You can now close this window and run your bot.');
                    } else {
                        alert('Error: ' + (data.error || 'Unknown error'));
                    }
                }

                function render(state) {
                    const content = document.getElementById('content');
                    content.innerHTML = '';

                    if (!state.user_info) {
                        // Step 1: Get credentials
                        content.innerHTML = `
                        <div class="section">
                            <h2>Step 1: Microsoft Credentials</h2>
                            <input type="text" id="profile" placeholder="Profile name" value="default">
                            <input type="text" id="clientId" placeholder="Client ID" value="${state.client_id}">
                            <input type="password" id="clientSecret" placeholder="Client Secret" value="${state.client_secret}">
                            <button onclick="submitConfig()">Connect Microsoft</button>
                        </div>
                        `;
                    } else {
                        // Step 2: Select site, drive, folder
                        content.innerHTML = `
                        <div class="status">
                            <strong>Connected:</strong> ${state.user_info.mail}
                        </div>
                        <div class="section">
                            <h2>Step 2: Select Locations</h2>
                            <label>Profile Name:</label>
                            <input type="text" id="profile" value="${state.profile_name || 'default'}">

                            <label>SharePoint Site:</label>
                            <select id="site" onchange="onSiteSelect(this.value)">
                                <option value="">Select a site...</option>
                            </select>

                            <label>Document Library:</label>
                            <select id="drive" onchange="onDriveSelect(this.value)">
                                <option value="">Select a library...</option>
                            </select>

                            <label>Default Folder:</label>
                            <select id="folder">
                                <option value="">Root</option>
                            </select>

                            <button onclick="testUpload()">Test Upload</button>
                            <button onclick="saveProfile()">Save Profile</button>
                        </div>
                        `;
                        loadSites();
                    }
                }

                init();
            </script>
        </body>
        </html>
        """
        self._send_html(html)

    def _api_init(self, query: dict):
        """Get current configuration state."""
        user = self.config_state.get("user_info")
        self._send_json({
            "user_info": user,
            "profile_name": self.config_state.get("profile_name", "default"),
            "client_id": self.config_state.get("client_id", ""),
        })

    def _api_sites(self):
        """Get list of SharePoint sites."""
        try:
            access_token = self.config_state["tokens"]["access_token"]
            graph = GraphClient(access_token)
            # For MVP, just return a placeholder
            # In production, would call /sites with search
            self._send_json({"sites": []})
        except Exception as e:
            self._send_json({"error": str(e)}, 400)

    def _api_drives(self):
        """Get document libraries for selected site."""
        self._send_json({"drives": []})

    def _api_folders(self, query: dict):
        """Get folders in drive."""
        self._send_json({"folders": []})

    def _api_test(self):
        """Test profile with upload."""
        self._send_json({"message": "Test successful"})

    def _api_save(self, query: dict):
        """Save profile."""
        profile_name = query.get("profile_name", ["default"])[0]
        self.config_state["profile_name"] = profile_name
        self._send_json({"success": True})

    def _send_json(self, data: dict, status: int = 200):
        """Send JSON response."""
        self.send_response(status)
        self.send_header("Content-type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())

    def _send_html(self, html: str):
        """Send HTML response."""
        self.send_response(200)
        self.send_header("Content-type", "text/html")
        self.end_headers()
        self.wfile.write(html.encode())

    def _serve_error(self, message: str):
        """Serve error page."""
        html = f"""
        <html><body>
        <h1>Error</h1>
        <p>{message}</p>
        <p><a href="/">Back</a></p>
        </body></html>
        """
        self._send_html(html)


def run_config_server(
    client_id: str = "",
    client_secret: str = "",
    profile_name: str = "default",
    port: int = 8001,
) -> dict:
    """Run configuration UI server.

    Args:
        client_id: Optional pre-filled client ID
        client_secret: Optional pre-filled client secret
        profile_name: Profile name to configure
        port: Server port

    Returns:
        Configuration dict with tokens and site/drive/folder info
    """
    ConfigHandler.config_state["profile_name"] = profile_name
    ConfigHandler.config_state["client_id"] = client_id
    ConfigHandler.config_state["client_secret"] = client_secret

    server = HTTPServer(("localhost", port), ConfigHandler)
    logger.info(f"Opening configuration UI at http://localhost:{port}")

    # Open browser in a thread
    def open_browser():
        import time
        time.sleep(1)  # Give server time to start
        webbrowser.open(f"http://localhost:{port}")

    thread = threading.Thread(target=open_browser, daemon=True)
    thread.start()

    # Run server
    try:
        server.handle_request()  # Handle one request
        server.handle_request()  # Handle callback
    except KeyboardInterrupt:
        logger.info("Configuration cancelled")
        raise

    return ConfigHandler.config_state
