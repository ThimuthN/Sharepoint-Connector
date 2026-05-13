"""Tests for SharePointClient SDK."""
import pytest
from unittest.mock import patch
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
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MockStore.return_value
            mock_store.load_profile.return_value = mock_profile

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth'):
                with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                    mock_graph = MockGraph.return_value
                    mock_graph.upload_file.return_value = {"id": "file_123"}

                    client = SharePointClient(profile="demo")
                    item_id = client.upload("local.pdf", "Invoices/local.pdf")

                    assert item_id == "file_123"
                    mock_graph.upload_file.assert_called_once_with(
                        "drive_456",
                        "local.pdf",
                        "Invoices/local.pdf",
                        conflict="overwrite",
                    )
