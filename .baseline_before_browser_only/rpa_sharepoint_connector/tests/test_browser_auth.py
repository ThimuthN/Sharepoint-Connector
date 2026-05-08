"""Tests for browser OAuth (Authorization Code + PKCE)."""
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch
from urllib.parse import parse_qs, urlparse

import pytest

from rpa_sharepoint_connector.browser_auth import (
    LocalOAuthCallbackServer,
    MicrosoftBrowserAuth,
)


class TestBrowserAuthUrl:
    """Test authorization URL generation with PKCE."""

    def test_configure_browser_builds_auth_url_with_state_and_code_challenge(self):
        """Auth URL must include state + PKCE challenge."""
        auth = MicrosoftBrowserAuth(
            client_id="client_123",
            tenant_id="common",
            redirect_uri="http://localhost:8765/callback",
        )

        request = auth.build_authorization_request()
        parsed = urlparse(request["authorization_url"])
        params = parse_qs(parsed.query)

        assert parsed.path.endswith("/authorize")
        assert params["state"][0] == request["state"]
        assert params["code_challenge"][0] == request["code_challenge"]
        assert params["code_challenge_method"][0] == "S256"
        assert params["client_id"][0] == "client_123"


class TestCallbackValidation:
    """Test local callback validation and shutdown behavior."""

    def test_callback_rejects_invalid_state(self):
        """Callback must reject state mismatch."""
        callback = LocalOAuthCallbackServer(
            expected_state="expected_state",
            redirect_uri="http://localhost:8765/callback",
        )

        status, message = callback.process_callback_params(
            {
                "state": "wrong_state",
                "code": "abc123",
            }
        )

        assert status == 400
        assert "state" in message.lower()
        assert callback.error is not None

    def test_callback_server_shuts_down(self):
        """Server must always close after callback wait completes."""
        class FakeHTTPServer:
            def __init__(self, address, handler):
                self.address = address
                self.handler = handler
                self.timeout = None
                self.closed = False

            def handle_request(self):
                return

            def server_close(self):
                self.closed = True

        callback = LocalOAuthCallbackServer(
            expected_state="expected_state",
            redirect_uri="http://localhost:8765/callback",
            server_class=FakeHTTPServer,
        )
        callback.error = "forced failure for test"

        with pytest.raises(ValueError, match="forced failure"):
            callback.wait_for_callback(timeout_seconds=2)

        assert callback.closed is True


class TestTokenExchange:
    """Test token exchange payload for PKCE public client."""

    def test_token_exchange_sends_code_verifier_without_client_secret(self):
        """Token exchange must send code_verifier and no client_secret."""
        auth = MicrosoftBrowserAuth(
            client_id="client_123",
            tenant_id="common",
            redirect_uri="http://localhost:8765/callback",
        )

        with patch("rpa_sharepoint_connector.browser_auth.httpx.Client") as MockClient:
            mock_response = MagicMock()
            mock_response.is_success = True
            mock_response.json.return_value = {
                "access_token": "access_token_123",
                "refresh_token": "refresh_token_456",
                "expires_in": 3600,
            }
            MockClient.return_value.__enter__.return_value.post.return_value = mock_response

            auth.exchange_code_for_tokens("auth_code_abc", "verifier_xyz")

            call_args = MockClient.return_value.__enter__.return_value.post.call_args
            data = call_args[1]["data"]
            assert data["code_verifier"] == "verifier_xyz"
            assert "client_secret" not in data


class TestRuntimeNoBrowser:
    """Runtime SDK must remain non-interactive (no browser)."""

    def test_runtime_sdk_never_opens_browser(self):
        """SharePointClient initialization should not open a browser."""
        profile_data = {
            "access_token": "access_token_123",
            "refresh_token": "refresh_token_456",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "site_id": "site_1",
            "site_name": "Site",
            "drive_id": "drive_1",
            "drive_name": "Documents",
            "folder_id": "root",
            "folder_path": "",
            "user_id": "user_1",
            "user_email": "user@example.com",
        }

        with patch("rpa_sharepoint_connector.sdk.TokenStore") as MockStore:
            with patch("rpa_sharepoint_connector.sdk.MicrosoftAuth") as MockAuth:
                with patch("rpa_sharepoint_connector.sdk.GraphClient"):
                    with patch("webbrowser.open") as mock_open:
                        mock_store = MagicMock()
                        mock_store.load_profile.return_value = profile_data
                        MockStore.return_value = mock_store

                        mock_auth = MagicMock()
                        MockAuth.return_value = mock_auth

                        from rpa_sharepoint_connector import SharePointClient
                        SharePointClient(profile="demo")

                        mock_open.assert_not_called()
