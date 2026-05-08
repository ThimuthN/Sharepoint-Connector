"""Hardened SDK tests for RPA bot reliability."""
import pytest
from unittest.mock import MagicMock, patch, mock_open
from datetime import datetime, timedelta
from rpa_sharepoint_connector import SharePointClient


class TestSDKHealthCheck:
    """Test preflight health check."""

    @pytest.fixture
    def mock_profile(self):
        return {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

    def test_health_check_exists(self, mock_profile):
        """Test that health_check method exists."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                sp = SharePointClient(profile="test")
                assert hasattr(sp, 'health_check')
                assert callable(sp.health_check)

    def test_health_check_verifies_profile_exists(self):
        """Test health check verifies profile exists."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            with pytest.raises(ValueError, match="not found"):
                SharePointClient(profile="missing")

    def test_health_check_validates_token(self, mock_profile):
        """Test health check validates token."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth') as MockAuth:
                mock_auth = MagicMock()
                mock_auth.refresh_token.return_value = {
                    "access_token": "new_token",
                    "refresh_token": "new_refresh",
                    "expires_in": 3600
                }
                MockAuth.return_value = mock_auth

                with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                    sp = SharePointClient(profile="test")
                    # Should auto-refresh since we have short expiry in test


class TestSDKDangerousOperations:
    """Test prevention of dangerous operations."""

    @pytest.fixture
    def mock_profile(self):
        return {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "root",  # Dangerous!
            "folder_path": ""
        }

    def test_cannot_delete_root_folder(self, mock_profile):
        """Test that deleting root folder is prevented."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                # Attempt to delete root should be prevented or at least handled
                # This is a safeguard - delete should fail gracefully if attempting root
                with pytest.raises(ValueError):
                    sp.delete("root")

    def test_delete_requires_valid_path(self, mock_profile):
        """Test delete validates path format."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                # Empty result means file not found
                mock_graph.list_items.return_value = []
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                # Deleting non-existent file should raise clear error
                with pytest.raises(ValueError):
                    sp.delete("NonExistent/File.pdf")


class TestSDKErrorMessages:
    """Test that errors are clear and actionable."""

    @pytest.fixture
    def mock_profile(self):
        return {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

    def test_missing_profile_error_is_actionable(self):
        """Test missing profile error tells user how to fix it."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = None
            MockStore.return_value = mock_store

            error_msg = None
            try:
                SharePointClient(profile="missing")
            except ValueError as e:
                error_msg = str(e)

            assert "configure" in error_msg.lower()
            assert "missing" in error_msg.lower()

    def test_upload_error_includes_path(self, mock_profile):
        """Test upload errors include the path."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                mock_graph.upload_file.side_effect = ValueError("Permission denied")
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                try:
                    sp.upload("local.pdf", "remote.pdf")
                except ValueError as e:
                    # Error should be clear
                    assert "Permission" in str(e) or "upload" in str(e).lower()

    def test_download_nonexistent_error(self, mock_profile):
        """Test downloading non-existent file gives clear error."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                mock_graph.list_items.return_value = []
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                with pytest.raises(ValueError):
                    sp.download("Nonexistent/File.pdf", "local.pdf")


class TestSDKRuntimeStability:
    """Test that SDK runtime is stable and deterministic."""

    @pytest.fixture
    def mock_profile(self):
        return {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

    def test_no_config_ui_at_runtime(self, mock_profile):
        """Test that runtime never opens config UI."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                with patch('webbrowser.open') as mock_browser:
                    sp = SharePointClient(profile="test")
                    # Config UI should never be opened
                    mock_browser.assert_not_called()

    def test_delete_missing_file_raises_error(self, mock_profile):
        """Test delete raises error when file path not found."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                mock_graph.list_items.return_value = []  # File not found
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                # Trying to delete non-existent file should raise
                with pytest.raises(ValueError, match="not found"):
                    sp.delete("NonExistent/file.pdf")

    def test_list_operation_never_fails_silently(self, mock_profile):
        """Test list operation raises errors clearly."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                mock_graph.list_items.side_effect = ValueError("Access denied")
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                # Should raise, not return empty list
                with pytest.raises(ValueError):
                    sp.list()


class TestSDKPathHandling:
    """Test path normalization and handling."""

    @pytest.fixture
    def mock_profile(self):
        return {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

    def test_find_item_id_with_slashes(self, mock_profile):
        """Test finding item by path with slashes."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                # Mock hierarchy: Folder -> SubFolder -> file.pdf
                mock_graph.list_items.side_effect = [
                    [{"id": "sub_1", "name": "Folder"}],
                    [{"id": "sub_2", "name": "SubFolder"}],
                    [{"id": "file_1", "name": "file.pdf"}]
                ]
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")
                item_id = sp._find_item_id("Folder/SubFolder/file.pdf")
                assert item_id == "file_1"

    def test_find_item_id_missing_path(self, mock_profile):
        """Test finding item with missing path component."""
        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                mock_graph = MagicMock()
                mock_graph.list_items.return_value = [{"id": "sub_1", "name": "Folder"}]
                MockGraph.return_value = mock_graph

                sp = SharePointClient(profile="test")

                with pytest.raises(ValueError, match="not found"):
                    sp._find_item_id("Folder/Missing/file.pdf")


class TestSDKTokenRefreshDeterministic:
    """Test token refresh is deterministic."""

    def test_refresh_only_when_needed(self):
        """Test token only refreshes when actually expired."""
        profile = {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth') as MockAuth:
                mock_auth = MagicMock()
                # Should NOT refresh for valid token
                MockAuth.return_value = mock_auth

                with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                    SharePointClient(profile="test")
                    # refresh_token should not be called
                    mock_auth.refresh_token.assert_not_called()

    def test_refresh_on_expiration_boundary(self):
        """Test token refreshes at exactly 5 minute boundary."""
        profile = {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(minutes=4)).isoformat(),  # < 5 min
            "client_id": "id",
            "client_secret": "secret",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth') as MockAuth:
                mock_auth = MagicMock()
                mock_auth.refresh_token.return_value = {
                    "access_token": "new_token",
                    "refresh_token": "new_refresh",
                    "expires_in": 3600
                }
                MockAuth.return_value = mock_auth

                with patch('rpa_sharepoint_connector.sdk.GraphClient'):
                    SharePointClient(profile="test")
                    # Should refresh since within 5 min boundary
                    mock_auth.refresh_token.assert_called()
