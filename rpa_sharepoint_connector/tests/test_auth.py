"""Tests for runtime auth refresh and user-info helpers."""
from datetime import datetime, timedelta
from unittest.mock import MagicMock, patch

import pytest

from rpa_sharepoint_connector.auth import MicrosoftAuth


class TestMicrosoftAuthInitialization:
    """Test auth client initialization."""

    def test_init_with_defaults(self, monkeypatch):
        monkeypatch.delenv("MICROSOFT_TENANT_ID", raising=False)
        auth = MicrosoftAuth()
        assert auth.client_id is not None
        assert auth.tenant_id == "common"
        assert "common" in auth.authority

    def test_init_with_custom_tenant(self):
        auth = MicrosoftAuth(tenant_id="specific_tenant")
        assert auth.tenant_id == "specific_tenant"
        assert "specific_tenant" in auth.authority

    def test_init_with_tenant_from_env(self, monkeypatch):
        monkeypatch.setenv("MICROSOFT_TENANT_ID", "consumers")
        auth = MicrosoftAuth()
        assert auth.tenant_id == "consumers"
        assert "consumers" in auth.authority


class TestTokenRefresh:
    """Test token refresh for public client flow."""

    def test_refresh_token_success(self):
        auth = MicrosoftAuth(client_id="test_id")

        with patch("rpa_sharepoint_connector.auth.httpx.Client") as mock_client:
            mock_response = MagicMock()
            mock_response.is_success = True
            mock_response.json.return_value = {
                "access_token": "new_token",
                "refresh_token": "new_refresh",
                "expires_in": 3600,
            }
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            result = auth.refresh_token("old_refresh")
            assert result["access_token"] == "new_token"

            call_args = mock_client.return_value.__enter__.return_value.post.call_args
            data = call_args[1]["data"]
            assert data["client_id"] == "test_id"
            assert data["refresh_token"] == "old_refresh"
            assert data["grant_type"] == "refresh_token"
            assert "client_secret" not in data

    def test_refresh_invalid_grant(self):
        auth = MicrosoftAuth()

        with patch("rpa_sharepoint_connector.auth.httpx.Client") as mock_client:
            mock_response = MagicMock()
            mock_response.is_success = False
            mock_response.json.return_value = {"error": "invalid_grant"}
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            with pytest.raises(ValueError, match="Refresh token expired or invalid"):
                auth.refresh_token("expired_token")


class TestUserInfo:
    """Test getting user info from Graph."""

    def test_get_user_info_success(self):
        auth = MicrosoftAuth()

        with patch("rpa_sharepoint_connector.auth.httpx.Client") as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {"id": "user_123", "mail": "user@example.com"}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            result = auth.get_user_info("token_abc")
            assert result["id"] == "user_123"

            call_args = mock_client.return_value.__enter__.return_value.get.call_args
            headers = call_args[1]["headers"]
            assert headers["Authorization"] == "Bearer token_abc"


class TestTokenExpiration:
    """Test token expiration checking."""

    def test_is_token_expired_when_past_expiry(self):
        auth = MicrosoftAuth()
        past = datetime.utcnow() - timedelta(hours=1)
        assert auth.is_token_expired(past) is True

    def test_is_token_expired_near_expiry(self):
        auth = MicrosoftAuth()
        near_future = datetime.utcnow() + timedelta(minutes=3)
        assert auth.is_token_expired(near_future) is True

    def test_is_token_not_expired_with_buffer(self):
        auth = MicrosoftAuth()
        future = datetime.utcnow() + timedelta(minutes=10)
        assert auth.is_token_expired(future) is False
