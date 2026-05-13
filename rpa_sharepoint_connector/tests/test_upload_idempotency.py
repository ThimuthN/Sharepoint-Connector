"""Tests for upload idempotency and conflict handling."""
import pytest
import tempfile
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta
from rpa_sharepoint_connector import SharePointClient
from rpa_sharepoint_connector.graph_client import GraphClient


class TestUploadConflictModes:
    """Test upload conflict behavior."""

    @pytest.fixture
    def temp_file(self):
        """Create a temporary file for testing."""
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".pdf") as f:
            f.write("test content")
            temp_path = f.name
        yield temp_path
        import os
        try:
            os.unlink(temp_path)
        except:
            pass

    def test_default_conflict_is_overwrite(self, temp_file):
        """Test default conflict mode is overwrite (backwards compatible)."""
        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "id": "file_1",
                "name": "test.pdf"
            }
            mock_client.return_value.put.return_value = mock_response

            with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                mock_ensure.return_value = "folder_1"

                graph = GraphClient("token")
                # Default should be overwrite, should NOT check if exists
                result = graph.upload_file("drive_1", temp_file, "Folder/test.pdf")

                assert result["id"] == "file_1"

    def test_upload_fail_if_exists_with_existing_file(self, temp_file):
        """Test fail_if_exists raises error when file exists."""
        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            mock_list.return_value = [
                {"id": "existing_1", "name": "test.pdf", "file": {}}
            ]

            with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                mock_ensure.return_value = "folder_1"

                graph = GraphClient("token")

                with pytest.raises(ValueError, match="File already exists"):
                    graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="fail_if_exists"
                    )

    def test_upload_fail_if_exists_with_new_file(self, temp_file):
        """Test fail_if_exists succeeds when file doesn't exist."""
        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            mock_list.return_value = [
                {"id": "other_1", "name": "other.pdf", "file": {}}
            ]

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                mock_response = MagicMock()
                mock_response.json.return_value = {
                    "id": "file_1",
                    "name": "test.pdf"
                }
                mock_client.return_value.put.return_value = mock_response

                with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                    mock_ensure.return_value = "folder_1"

                    graph = GraphClient("token")
                    result = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="fail_if_exists"
                    )

                    assert result["id"] == "file_1"

    def test_upload_overwrite_replaces_existing(self, temp_file):
        """Test overwrite mode replaces existing file."""
        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.json.return_value = {
                "id": "file_1",
                "name": "test.pdf"
            }
            mock_client.return_value.put.return_value = mock_response

            with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                mock_ensure.return_value = "folder_1"

                graph = GraphClient("token")
                # Overwrite should NOT check if exists
                result = graph.upload_file(
                    "drive_1", temp_file, "Folder/test.pdf",
                    conflict="overwrite"
                )

                assert result["id"] == "file_1"

    def test_upload_rename_creates_unique_name(self, temp_file):
        """Test rename mode creates unique filename when exists."""
        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            mock_list.return_value = [
                {"id": "existing_1", "name": "test.pdf", "file": {}}
            ]

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                mock_response = MagicMock()
                mock_response.json.return_value = {
                    "id": "file_2",
                    "name": "test (1).pdf"
                }
                mock_client.return_value.put.return_value = mock_response

                with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                    mock_ensure.return_value = "folder_1"

                    graph = GraphClient("token")
                    result = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="rename"
                    )

                    # Check that rename happened
                    call_url = mock_client.return_value.put.call_args[0][0]
                    assert "test (1).pdf" in call_url

    def test_upload_rename_multiple_conflicts(self, temp_file):
        """Test rename mode with multiple conflicting files."""
        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            mock_list.return_value = [
                {"id": "existing_1", "name": "test.pdf", "file": {}},
                {"id": "existing_2", "name": "test (1).pdf", "file": {}},
                {"id": "existing_3", "name": "test (2).pdf", "file": {}}
            ]

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                mock_response = MagicMock()
                mock_response.json.return_value = {
                    "id": "file_4",
                    "name": "test (3).pdf"
                }
                mock_client.return_value.put.return_value = mock_response

                with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                    mock_ensure.return_value = "folder_1"

                    graph = GraphClient("token")
                    result = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="rename"
                    )

                    # Should use test (3).pdf (next available)
                    call_url = mock_client.return_value.put.call_args[0][0]
                    assert "test (3).pdf" in call_url

    def test_upload_invalid_conflict_mode(self, temp_file):
        """Test invalid conflict mode raises error."""
        graph = GraphClient("token")

        with pytest.raises(ValueError, match="Invalid conflict mode"):
            graph.upload_file(
                "drive_1", temp_file, "Folder/test.pdf",
                conflict="invalid_mode"
            )


class TestGenerateUniqueFilename:
    """Test unique filename generation."""

    def test_no_conflict(self):
        """Test filename unchanged when no conflict."""
        graph = GraphClient("token")
        existing = {"other.pdf", "file.txt"}

        result = graph._generate_unique_filename("test.pdf", existing)
        assert result == "test.pdf"

    def test_single_conflict(self):
        """Test generates (1) suffix on conflict."""
        graph = GraphClient("token")
        existing = {"test.pdf", "other.txt"}

        result = graph._generate_unique_filename("test.pdf", existing)
        assert result == "test (1).pdf"

    def test_multiple_conflicts(self):
        """Test finds next available number."""
        graph = GraphClient("token")
        existing = {"test.pdf", "test (1).pdf", "test (2).pdf"}

        result = graph._generate_unique_filename("test.pdf", existing)
        assert result == "test (3).pdf"

    def test_no_extension(self):
        """Test works with files without extension."""
        graph = GraphClient("token")

        # Without conflict
        existing = {"other.txt"}
        result = graph._generate_unique_filename("README", existing)
        assert result == "README"

        # With conflict
        existing = {"README", "other.txt"}
        result = graph._generate_unique_filename("README", existing)
        assert result == "README (1)"

    def test_multiple_dots_in_name(self):
        """Test handles filenames with multiple dots."""
        graph = GraphClient("token")
        existing = {"report.2024.pdf"}

        result = graph._generate_unique_filename("report.2024.pdf", existing)
        assert result == "report.2024 (1).pdf"


class TestSDKUploadConflict:
    """Test SDK-level upload conflict parameter."""

    @pytest.fixture
    def temp_file(self):
        """Create a temporary file for testing."""
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".pdf") as f:
            f.write("test content")
            temp_path = f.name
        yield temp_path
        import os
        try:
            os.unlink(temp_path)
        except:
            pass

    def test_sdk_upload_passes_conflict_param(self, temp_file):
        """Test SDK upload passes conflict parameter to GraphClient."""
        mock_profile = {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth'):
                with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                    mock_graph = MagicMock()
                    mock_graph.upload_file.return_value = {"id": "file_1"}
                    MockGraph.return_value = mock_graph

                    sp = SharePointClient(profile="test")

                    # Call upload with conflict parameter
                    sp.upload(temp_file, "test.pdf", conflict="fail_if_exists")

                    # Verify conflict parameter was passed
                    mock_graph.upload_file.assert_called_once()
                    call_args = mock_graph.upload_file.call_args
                    assert call_args[1]["conflict"] == "fail_if_exists"

    def test_sdk_upload_default_conflict(self, temp_file):
        """Test SDK upload defaults to overwrite for backwards compatibility."""
        mock_profile = {
            "access_token": "token",
            "refresh_token": "refresh",
            "expires_at": (datetime.utcnow() + timedelta(hours=1)).isoformat(),
            "client_id": "id",
            "site_id": "site_1",
            "drive_id": "drive_1",
            "folder_id": "folder_1",
            "folder_path": "Invoices"
        }

        with patch('rpa_sharepoint_connector.sdk.TokenStore') as MockStore:
            mock_store = MagicMock()
            mock_store.load_profile.return_value = mock_profile
            MockStore.return_value = mock_store

            with patch('rpa_sharepoint_connector.sdk.MicrosoftAuth'):
                with patch('rpa_sharepoint_connector.sdk.GraphClient') as MockGraph:
                    mock_graph = MagicMock()
                    mock_graph.upload_file.return_value = {"id": "file_1"}
                    MockGraph.return_value = mock_graph

                    sp = SharePointClient(profile="test")

                    # Call upload without conflict parameter
                    sp.upload(temp_file, "test.pdf")

                    # Verify defaults to overwrite
                    call_args = mock_graph.upload_file.call_args
                    assert call_args[1]["conflict"] == "overwrite"


class TestRetryIdempotency:
    """Test retry safety with conflict modes."""

    @pytest.fixture
    def temp_file(self):
        """Create a temporary file for testing."""
        with tempfile.NamedTemporaryFile(mode="w", delete=False, suffix=".pdf") as f:
            f.write("test content")
            temp_path = f.name
        yield temp_path
        import os
        try:
            os.unlink(temp_path)
        except:
            pass

    def test_fail_if_exists_prevents_retry_duplicate(self, temp_file):
        """Test fail_if_exists prevents duplicate after network failure."""
        # Scenario: upload succeeds, network dies before response,
        # bot retries, fail_if_exists detects file and raises error

        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            # First call (pre-upload check): file doesn't exist
            # This represents the first retry attempt
            mock_list.side_effect = [
                [{"id": "other", "name": "other.pdf"}],  # Pre-upload check: no file
                [{"id": "file_1", "name": "test.pdf", "file": {}}]  # Retry: file exists
            ]

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                # First upload succeeds but response lost
                # Retry attempt finds file exists
                mock_response = MagicMock()
                mock_response.json.return_value = {"id": "file_1", "name": "test.pdf"}
                mock_client.return_value.put.return_value = mock_response

                with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                    mock_ensure.return_value = "folder_1"

                    graph = GraphClient("token")

                    # First call succeeds
                    result = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="fail_if_exists"
                    )
                    assert result["id"] == "file_1"

                    # Retry would detect file exists
                    with pytest.raises(ValueError, match="File already exists"):
                        graph.upload_file(
                            "drive_1", temp_file, "Folder/test.pdf",
                            conflict="fail_if_exists"
                        )

    def test_rename_handles_concurrent_uploads(self, temp_file):
        """Test rename mode handles concurrent uploads creating duplicates."""
        # Scenario: two bots both rename to (1), second one detects conflict

        with patch('rpa_sharepoint_connector.graph_client.GraphClient.list_items') as mock_list:
            # First call: only original exists
            # Second call: now both original and (1) exist
            mock_list.side_effect = [
                [{"name": "test.pdf", "file": {}}],
                [{"name": "test.pdf", "file": {}}, {"name": "test (1).pdf", "file": {}}]
            ]

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                mock_response1 = MagicMock()
                mock_response1.json.return_value = {"id": "file_1", "name": "test (1).pdf"}

                mock_response2 = MagicMock()
                mock_response2.json.return_value = {"id": "file_2", "name": "test (2).pdf"}

                mock_client.return_value.put.side_effect = [
                    mock_response1, mock_response2
                ]

                with patch('rpa_sharepoint_connector.graph_client.GraphClient._ensure_folder_path') as mock_ensure:
                    mock_ensure.return_value = "folder_1"

                    graph = GraphClient("token")

                    # First bot uploads as test (1).pdf
                    result1 = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="rename"
                    )

                    # Second bot retries, finds (1) exists, creates (2).pdf
                    result2 = graph.upload_file(
                        "drive_1", temp_file, "Folder/test.pdf",
                        conflict="rename"
                    )

                    assert result1["name"] == "test (1).pdf"
                    assert result2["name"] == "test (2).pdf"
