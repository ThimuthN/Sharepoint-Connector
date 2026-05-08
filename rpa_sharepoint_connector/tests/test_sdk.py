"""Tests for SharePointClient SDK."""
import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta


class TestSharePointClient:
    """Test SharePointClient operations."""

    @pytest.fixture
    def mock_profile(self):
        """Mock profile data."""
        return {
            "access_token": "test_token_123",
            "refresh_token": "test_refresh_456",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "site_id": "site_123",
            "site_name": "Test Site",
            "drive_id": "drive_456",
            "drive_name": "Documents",
            "folder_id": "folder_789",
            "folder_path": "Invoices",
        }

    def test_client_init_missing_profile(self):
        """Test client init fails with missing profile."""
        from rpa_sharepoint_connector import SharePointClient
        with pytest.raises(ValueError, match="not found"):
            SharePointClient(profile="nonexistent")

    def test_upload_succeeds(self, mock_profile):
        """Test upload operation."""
        from rpa_sharepoint_connector import SharePointClient
        with patch('rpa_sharepoint_connector.sdk.TokenStore'):
            with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                pass  # Would test with mocks
