"""Tests for Device Code Flow authentication (public client, no secrets)."""
import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta
from rpa_sharepoint_connector.auth import MicrosoftAuth


class TestMicrosoftAuthInitialization:
    """Test auth client initialization."""

    def test_init_with_defaults(self, monkeypatch):
        """Test initialization with default Innobot app."""
        monkeypatch.delenv("MICROSOFT_TENANT_ID", raising=False)
        auth = MicrosoftAuth()
        assert auth.client_id is not None
        assert auth.tenant_id == "common"
        assert "common" in auth.authority
        assert "offline_access" in auth.scopes
        assert "User.Read" in auth.scopes
        assert "Files.ReadWrite.All" in auth.scopes
        assert "Sites.ReadWrite.All" in auth.scopes

    def test_init_with_custom_client_id(self):
        """Test initialization with custom client ID."""
        auth = MicrosoftAuth(client_id="custom_id")
        assert auth.client_id == "custom_id"

    def test_init_with_custom_tenant(self):
        """Test initialization with custom tenant."""
        auth = MicrosoftAuth(tenant_id="specific_tenant")
        assert auth.tenant_id == "specific_tenant"
        assert "specific_tenant" in auth.authority

    def test_init_with_tenant_from_env(self, monkeypatch):
        """Test initialization reads tenant from MICROSOFT_TENANT_ID."""
        monkeypatch.setenv("MICROSOFT_TENANT_ID", "consumers")
        auth = MicrosoftAuth()
        assert auth.tenant_id == "consumers"
        assert "consumers" in auth.authority

    def test_no_client_secret_exists(self):
        """Test that no client secret is stored (public client)."""
        auth = MicrosoftAuth()
        assert not hasattr(auth, "client_secret")


class TestDeviceCodeFlow:
    """Test Device Code Flow initiation."""

    def test_start_device_flow_success(self):
        """Test successful device flow initiation."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "device_code": "dev_code_123",
                "user_code": "ABC-DEF",
                "verification_uri": "https://microsoft.com/devicelogin",
                "verification_uri_complete": "https://microsoft.com/devicelogin?user_code=ABC-DEF",
                "expires_in": 900,
                "interval": 5,
                "message": "To sign in, use a web browser to open the page..."
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            result = auth.start_device_flow()

            assert result["device_code"] == "dev_code_123"
            assert result["user_code"] == "ABC-DEF"
            assert "verification_uri" in result
            assert result["expires_in"] == 900

    def test_start_device_flow_sends_client_id_only(self):
        """Test device flow request uses client_id (no secret)."""
        auth = MicrosoftAuth(client_id="test_id")

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "device_code": "dev_code",
                "user_code": "ABC",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            auth.start_device_flow()

            call_args = mock_client.return_value.__enter__.return_value.post.call_args
            data = call_args[1]["data"]
            assert data["client_id"] == "test_id"
            assert "client_secret" not in data

    def test_start_device_flow_includes_scopes(self):
        """Test device flow request includes required scopes."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "device_code": "dev_code",
                "user_code": "ABC",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            auth.start_device_flow()

            call_args = mock_client.return_value.__enter__.return_value.post.call_args
            data = call_args[1]["data"]
            assert "offline_access" in data["scope"]
            assert "Files.ReadWrite.All" in data["scope"]

    def test_start_device_flow_failure(self):
        """Test device flow initiation failure."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            import httpx
            mock_client.return_value.__enter__.return_value.post.side_effect = httpx.HTTPError(
                "Network error"
            )

            with pytest.raises(ValueError, match="Failed to start device flow"):
                auth.start_device_flow()


class TestDeviceCodePolling:
    """Test Device Code Flow polling."""

    def test_poll_authorization_pending(self):
        """Test polling handles authorization_pending."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.sleep'):
                with patch('rpa_sharepoint_connector.auth.time.time') as mock_time:
                    # Simulate: first call returns pending, second returns success
                    mock_time.side_effect = [0, 0.1, 1]

                    mock_response1 = MagicMock()
                    mock_response1.is_success = False
                    mock_response1.json.return_value = {"error": "authorization_pending"}

                    mock_response2 = MagicMock()
                    mock_response2.is_success = True
                    mock_response2.json.return_value = {
                        "access_token": "token_123",
                        "refresh_token": "refresh_456",
                        "expires_in": 3600
                    }

                    mock_client.return_value.__enter__.return_value.post.side_effect = [
                        mock_response1,
                        mock_response2
                    ]

                    result = auth.poll_device_flow("dev_code", 5, 900)

                    assert result["access_token"] == "token_123"
                    assert result["refresh_token"] == "refresh_456"

    def test_poll_slow_down(self):
        """Test polling handles slow_down instruction."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.sleep'):
                with patch('rpa_sharepoint_connector.auth.time.time') as mock_time:
                    mock_time.side_effect = [0, 0.1, 0.2, 1]

                    mock_response1 = MagicMock()
                    mock_response1.is_success = False
                    mock_response1.json.return_value = {"error": "slow_down"}

                    mock_response2 = MagicMock()
                    mock_response2.is_success = True
                    mock_response2.json.return_value = {
                        "access_token": "token",
                        "refresh_token": "refresh",
                        "expires_in": 3600
                    }

                    mock_client.return_value.__enter__.return_value.post.side_effect = [
                        mock_response1,
                        mock_response2
                    ]

                    result = auth.poll_device_flow("dev_code", 5, 900)
                    assert result["access_token"] == "token"

    def test_poll_expired_token(self):
        """Test polling raises error on expired_token."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.time', return_value=0):
                mock_response = MagicMock()
                mock_response.is_success = False
                mock_response.json.return_value = {"error": "expired_token"}
                mock_client.return_value.__enter__.return_value.post.return_value = mock_response

                with pytest.raises(ValueError, match="Device code expired"):
                    auth.poll_device_flow("dev_code", 5, 900)

    def test_poll_access_denied(self):
        """Test polling raises error on access_denied."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.time', return_value=0):
                mock_response = MagicMock()
                mock_response.is_success = False
                mock_response.json.return_value = {"error": "access_denied"}
                mock_client.return_value.__enter__.return_value.post.return_value = mock_response

                with pytest.raises(ValueError, match="Authorization denied"):
                    auth.poll_device_flow("dev_code", 5, 900)

    def test_poll_timeout(self):
        """Test polling raises error on timeout."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.sleep'):
                with patch('rpa_sharepoint_connector.auth.time.time') as mock_time:
                    # Simulate time passing beyond expires_in
                    mock_time.side_effect = [0, 1000]

                    mock_response = MagicMock()
                    mock_response.is_success = False
                    mock_response.json.return_value = {"error": "authorization_pending"}
                    mock_client.return_value.__enter__.return_value.post.return_value = mock_response

                    with pytest.raises(ValueError, match="timed out"):
                        auth.poll_device_flow("dev_code", 5, 900)

    def test_poll_success(self):
        """Test successful polling and token retrieval."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            with patch('rpa_sharepoint_connector.auth.time.time', return_value=0):
                mock_response = MagicMock()
                mock_response.is_success = True
                mock_response.json.return_value = {
                    "access_token": "access_123",
                    "refresh_token": "refresh_456",
                    "expires_in": 3600,
                    "token_type": "Bearer"
                }
                mock_client.return_value.__enter__.return_value.post.return_value = mock_response

                result = auth.poll_device_flow("dev_code", 5, 900)

                assert result["access_token"] == "access_123"
                assert result["refresh_token"] == "refresh_456"


class TestTokenRefresh:
    """Test token refresh (public client, no secret)."""

    def test_refresh_token_success(self):
        """Test successful token refresh."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.is_success = True
            mock_response.json.return_value = {
                "access_token": "new_token",
                "refresh_token": "new_refresh",
                "expires_in": 3600
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            result = auth.refresh_token("old_refresh")

            assert result["access_token"] == "new_token"
            assert result["refresh_token"] == "new_refresh"

    def test_refresh_token_no_secret(self):
        """Test refresh request does not include client_secret."""
        auth = MicrosoftAuth(client_id="test_id")

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.is_success = True
            mock_response.json.return_value = {
                "access_token": "token",
                "refresh_token": "refresh",
                "expires_in": 3600
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            auth.refresh_token("refresh_token")

            call_args = mock_client.return_value.__enter__.return_value.post.call_args
            data = call_args[1]["data"]
            assert data["client_id"] == "test_id"
            assert data["refresh_token"] == "refresh_token"
            assert "client_secret" not in data

    def test_refresh_invalid_grant(self):
        """Test invalid_grant error is caught and provides clear message."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.is_success = False
            mock_response.json.return_value = {"error": "invalid_grant"}
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            with pytest.raises(ValueError, match="Refresh token expired or invalid"):
                auth.refresh_token("expired_token")

    def test_refresh_other_error(self):
        """Test other refresh errors raise ValueError."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            import httpx
            mock_client.return_value.__enter__.return_value.post.side_effect = httpx.HTTPError(
                "Server error"
            )

            with pytest.raises(ValueError, match="Failed to refresh"):
                auth.refresh_token("token")


class TestUserInfo:
    """Test getting user info."""

    def test_get_user_info_success(self):
        """Test successful user info fetch."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "id": "user_123",
                "mail": "user@example.com"
            }
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            result = auth.get_user_info("token_abc")

            assert result["id"] == "user_123"
            assert result["mail"] == "user@example.com"

    def test_get_user_info_includes_auth_header(self):
        """Test that auth header is sent."""
        auth = MicrosoftAuth()

        with patch('rpa_sharepoint_connector.auth.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {"id": "user_123"}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            auth.get_user_info("token_xyz")

            call_args = mock_client.return_value.__enter__.return_value.get.call_args
            headers = call_args[1]["headers"]
            assert "Authorization" in headers
            assert "Bearer token_xyz" in headers["Authorization"]


class TestTokenExpiration:
    """Test token expiration checking."""

    def test_is_token_expired_when_past_expiry(self):
        """Test expired token detection."""
        auth = MicrosoftAuth()
        past = datetime.utcnow() - timedelta(hours=1)

        assert auth.is_token_expired(past) is True

    def test_is_token_expired_near_expiry(self):
        """Test token near expiry (5 min buffer)."""
        auth = MicrosoftAuth()
        near_future = datetime.utcnow() + timedelta(minutes=3)

        assert auth.is_token_expired(near_future) is True

    def test_is_token_not_expired_with_buffer(self):
        """Test token with sufficient time remaining."""
        auth = MicrosoftAuth()
        future = datetime.utcnow() + timedelta(minutes=10)

        assert auth.is_token_expired(future) is False
