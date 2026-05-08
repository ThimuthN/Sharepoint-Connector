"""Tests for browser OAuth (Authorization Code + PKCE)."""
from datetime import datetime, timedelta
import hashlib
import time
import base64
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

    def test_pkce_challenge_matches_verifier(self):
        """PKCE challenge must match SHA256(verifier)."""
        verifier = "test_verifier_abc123"
        challenge = MicrosoftBrowserAuth._create_code_challenge(verifier)
        expected = base64.urlsafe_b64encode(
            hashlib.sha256(verifier.encode("utf-8")).digest()
        ).decode("utf-8").rstrip("=")
        assert challenge == expected

    def test_configure_opens_browser_auth_url(self):
        """Interactive auth should open the generated authorization URL."""
        auth = MicrosoftBrowserAuth(
            client_id="client_123",
            tenant_id="common",
            redirect_uri="http://localhost:8765/callback",
        )
        request = auth.build_authorization_request()
        opened = []

        class FakeCallbackServer:
            def __init__(self, expected_state, redirect_uri):
                self.expected_state = expected_state
                self.redirect_uri = redirect_uri
                self.closed = True

            def wait_for_callback(self, timeout_seconds=900):
                # Give the browser-opener thread enough time to run.
                time.sleep(0.35)
                return "auth_code_123"

        with patch.object(auth, "exchange_code_for_tokens") as mock_exchange:
            with patch.object(auth, "get_user_info") as mock_user:
                mock_exchange.return_value = {
                    "access_token": "access_token_123",
                    "refresh_token": "refresh_token_456",
                    "expires_in": 3600,
                }
                mock_user.return_value = {"id": "user_1", "mail": "user@example.com"}

                auth.authenticate(
                    open_browser=True,
                    browser_opener=lambda url, new=1: opened.append(url),
                    callback_server_class=FakeCallbackServer,
                    authorization_request=request,
                )

        assert opened == [request["authorization_url"]]


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

    def test_future_runtime_loads_saved_refresh_token(self):
        """Future runtime should load saved profile tokens without re-auth."""
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
                with patch("rpa_sharepoint_connector.sdk.GraphClient") as MockGraph:
                    mock_store = MagicMock()
                    mock_store.load_profile.return_value = profile_data
                    MockStore.return_value = mock_store

                    mock_auth = MagicMock()
                    MockAuth.return_value = mock_auth

                    from rpa_sharepoint_connector import SharePointClient
                    client = SharePointClient(profile="demo")

                    assert client.profile_data["refresh_token"] == "refresh_token_456"
                    MockGraph.assert_called_once_with("access_token_123")

    def test_expired_access_token_refreshes_silently(self):
        """Expired runtime token should refresh without UI/browser."""
        profile_data = {
            "access_token": "old_access_token",
            "refresh_token": "refresh_token_456",
            "expires_at": (datetime.utcnow() - timedelta(minutes=10)).isoformat(),
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
                        mock_auth.refresh_token.return_value = {
                            "access_token": "new_access_token",
                            "refresh_token": "new_refresh_token",
                            "expires_in": 3600,
                        }
                        MockAuth.return_value = mock_auth

                        from rpa_sharepoint_connector import SharePointClient
                        client = SharePointClient(profile="demo")

                        assert client.profile_data["access_token"] == "new_access_token"
                        mock_auth.refresh_token.assert_called_once_with("refresh_token_456")
                        mock_store.save_profile.assert_called_once()
                        mock_open.assert_not_called()

    def test_missing_profile_error_mentions_configure(self):
        """Missing profile should tell user to run configure first."""
        with patch("rpa_sharepoint_connector.sdk.TokenStore") as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            from rpa_sharepoint_connector import SharePointClient
            with pytest.raises(ValueError, match="Run: python -m rpa_sharepoint_connector configure --profile"):
                SharePointClient(profile="missing_profile")
