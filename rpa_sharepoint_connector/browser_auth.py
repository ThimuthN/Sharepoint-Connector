"""Browser OAuth with Authorization Code Flow + PKCE for local CLI setup."""
import base64
import hashlib
import logging
import os
import secrets
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Dict, Optional, Tuple, Type
from urllib.parse import parse_qs, urlencode, urlparse

import httpx

from .auth import MICROSOFT_CLIENT_ID

logger = logging.getLogger(__name__)

DEFAULT_REDIRECT_URI = os.getenv(
    "MICROSOFT_BROWSER_REDIRECT_URI",
    "http://localhost:8765/callback",
)


class LocalOAuthCallbackServer:
    """Temporary local HTTP server for OAuth callback handling."""

    def __init__(
        self,
        expected_state: str,
        redirect_uri: str = DEFAULT_REDIRECT_URI,
        server_class: Type[HTTPServer] = HTTPServer,
    ):
        parsed = urlparse(redirect_uri)
        if parsed.scheme != "http":
            raise ValueError("Redirect URI must use http for local callback server.")
        if parsed.hostname not in ("localhost", "127.0.0.1"):
            raise ValueError("Redirect URI host must be localhost or 127.0.0.1.")
        if not parsed.port:
            raise ValueError("Redirect URI must include an explicit port.")

        self.expected_state = expected_state
        self.host = parsed.hostname
        self.port = parsed.port
        self.callback_path = parsed.path or "/callback"
        self.server_class = server_class
        self.received_code: Optional[str] = None
        self.received_state: Optional[str] = None
        self.error: Optional[str] = None
        self._closed = False

    @property
    def closed(self) -> bool:
        """Whether the callback server has been closed."""
        return self._closed

    def process_callback_params(self, query: Dict[str, str]) -> Tuple[int, str]:
        """Validate callback query parameters."""
        oauth_error = query.get("error", "")
        if oauth_error:
            self.error = f"Microsoft returned error: {oauth_error}"
            return 400, self.error

        state = query.get("state", "")
        self.received_state = state
        if state != self.expected_state:
            self.error = "Invalid state parameter."
            return 400, self.error

        code = query.get("code", "")
        if not code:
            self.error = "No authorization code received."
            return 400, self.error

        self.received_code = code
        return 200, "Authentication successful. You can close this window."

    def _build_handler(self):
        parent = self

        class CallbackHandler(BaseHTTPRequestHandler):
            def log_message(self, format, *args):
                """Silence default HTTP request logging."""
                return

            def do_GET(self):
                parsed = urlparse(self.path)
                if parsed.path != parent.callback_path:
                    self.send_response(404)
                    self.send_header("Content-type", "text/html")
                    self.end_headers()
                    self.wfile.write(b"<h1>Not Found</h1>")
                    return

                raw_params = parse_qs(parsed.query)
                query = {key: values[0] if values else "" for key, values in raw_params.items()}
                status, message = parent.process_callback_params(query)

                self.send_response(status)
                self.send_header("Content-type", "text/html")
                self.end_headers()
                body = f"<html><body><p>{message}</p></body></html>".encode()
                self.wfile.write(body)

        return CallbackHandler

    def wait_for_callback(self, timeout_seconds: int = 900) -> str:
        """Wait for OAuth callback and return authorization code."""
        try:
            server = self.server_class((self.host, self.port), self._build_handler())
        except OSError as exc:
            raise ValueError(
                f"Failed to start callback server on {self.host}:{self.port}. "
                "Check whether the port is already in use."
            ) from exc

        server.timeout = 1
        deadline = time.time() + timeout_seconds

        try:
            while time.time() < deadline and not self.received_code and not self.error:
                server.handle_request()

            if self.received_code:
                return self.received_code
            if self.error:
                raise ValueError(self.error)
            raise TimeoutError("OAuth callback timed out.")
        finally:
            server.server_close()
            self._closed = True


class MicrosoftBrowserAuth:
    """Authorization Code Flow with PKCE for interactive browser login."""

    def __init__(
        self,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        redirect_uri: Optional[str] = None,
    ):
        self.client_id = client_id or MICROSOFT_CLIENT_ID
        self.tenant_id = tenant_id or os.getenv("MICROSOFT_TENANT_ID", "common")
        self.redirect_uri = redirect_uri or DEFAULT_REDIRECT_URI
        self.scopes = [
            "offline_access",
            "User.Read",
            "Sites.ReadWrite.All",
            "Files.ReadWrite.All",
        ]
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0"

    @staticmethod
    def _generate_code_verifier() -> str:
        """Generate a high-entropy PKCE code verifier."""
        raw = secrets.token_urlsafe(64)
        return raw[:128]

    @staticmethod
    def _create_code_challenge(code_verifier: str) -> str:
        """Create PKCE code challenge from verifier."""
        digest = hashlib.sha256(code_verifier.encode("utf-8")).digest()
        return base64.urlsafe_b64encode(digest).decode("utf-8").rstrip("=")

    def build_authorization_request(self) -> Dict[str, str]:
        """Build authorization URL, state, and PKCE verifier."""
        state = secrets.token_urlsafe(32)
        code_verifier = self._generate_code_verifier()
        code_challenge = self._create_code_challenge(code_verifier)

        params = {
            "client_id": self.client_id,
            "redirect_uri": self.redirect_uri,
            "response_type": "code",
            "response_mode": "query",
            "scope": " ".join(self.scopes),
            "state": state,
            "code_challenge": code_challenge,
            "code_challenge_method": "S256",
            "prompt": "select_account",
        }

        authorization_url = f"{self.authority}/authorize?{urlencode(params)}"
        logger.info(
            "Built browser authorization request: tenant=%s redirect_uri=%s",
            self.tenant_id,
            self.redirect_uri,
        )
        return {
            "state": state,
            "code_verifier": code_verifier,
            "code_challenge": code_challenge,
            "authorization_url": authorization_url,
        }

    def exchange_code_for_tokens(self, code: str, code_verifier: str) -> Dict:
        """Exchange authorization code for tokens using PKCE."""
        data = {
            "client_id": self.client_id,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": self.redirect_uri,
            "code_verifier": code_verifier,
            "scope": " ".join(self.scopes),
        }

        token_url = f"{self.authority}/token"
        try:
            with httpx.Client() as client:
                response = client.post(token_url, data=data, timeout=15.0)
                if not response.is_success:
                    try:
                        error_data = response.json()
                        error = error_data.get("error", "unknown_error")
                        description = error_data.get("error_description", "")
                        raise ValueError(
                            f"Token exchange failed: {error}. {description}".strip()
                        )
                    except ValueError:
                        raise
                    except Exception:
                        response.raise_for_status()
                return response.json()
        except httpx.HTTPError as exc:
            raise ValueError(f"Failed to exchange code for tokens: {exc}") from exc

    def get_user_info(self, access_token: str) -> Dict:
        """Get current Microsoft user profile."""
        try:
            with httpx.Client() as client:
                response = client.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {access_token}"},
                    timeout=15.0,
                )
                response.raise_for_status()
                return response.json()
        except httpx.HTTPError as exc:
            raise ValueError(f"Failed to get user info: {exc}") from exc

    def authenticate(
        self,
        timeout_seconds: int = 900,
        open_browser: bool = True,
        browser_opener=webbrowser.open,
        callback_server_class=LocalOAuthCallbackServer,
        authorization_request: Optional[Dict[str, str]] = None,
    ) -> Dict:
        """Run interactive browser authentication and return tokens + user info."""
        request = authorization_request or self.build_authorization_request()
        callback_server = callback_server_class(
            expected_state=request["state"],
            redirect_uri=self.redirect_uri,
        )

        if open_browser:
            def _open_browser():
                # Ensure callback server bind happens before browser redirects back.
                time.sleep(0.25)
                browser_opener(request["authorization_url"], new=1)

            threading.Thread(target=_open_browser, daemon=True).start()

        code = callback_server.wait_for_callback(timeout_seconds=timeout_seconds)
        tokens = self.exchange_code_for_tokens(code, request["code_verifier"])
        user_info = self.get_user_info(tokens["access_token"])
        return {
            "tokens": tokens,
            "user_info": user_info,
            "authorization_url": request["authorization_url"],
            "state": request["state"],
            "code_challenge": request["code_challenge"],
            "callback_closed": callback_server.closed,
        }
