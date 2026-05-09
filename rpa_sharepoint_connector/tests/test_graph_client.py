"""Tests for Microsoft Graph API client."""
import pytest
from unittest.mock import MagicMock, patch
import httpx
from rpa_sharepoint_connector.graph_client import GraphClient


class TestGraphClientInitialization:
    """Test graph client setup."""

    def test_init_with_token(self):
        """Test client initialization."""
        client = GraphClient("token_123")
        assert client.access_token == "token_123"
        assert "Bearer token_123" in client.headers["Authorization"]


class TestGraphClientGetMe:
    """Test getting user info."""

    def test_get_me_success(self):
        """Test successful /me call."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"id": "user_123", "mail": "user@example.com"}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            result = client.get_me()
            assert result["id"] == "user_123"


class TestGraphClientListDrives:
    """Test listing drives."""

    def test_list_drives_success(self):
        """Test successful drive listing."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {
                "value": [
                    {"id": "drive_1", "name": "Documents"},
                    {"id": "drive_2", "name": "Archive"}
                ]
            }
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            drives = client.list_drives("site_123")
            assert len(drives) == 2
            assert drives[0]["name"] == "Documents"

    def test_list_drives_empty(self):
        """Test empty drive list."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"value": []}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            drives = client.list_drives("site_123")
            assert len(drives) == 0


class TestGraphClientUpload:
    """Test file upload."""

    def test_upload_success(self):
        """Test successful file upload."""
        client = GraphClient("token")

        with patch('builtins.open', create=True) as mock_open:
            mock_open.return_value.__enter__.return_value.read.return_value = b"file content"

            with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                mock_response = MagicMock()
                mock_response.status_code = 201
                mock_response.json.return_value = {"id": "file_123", "name": "test.pdf"}
                mock_client.return_value.__enter__.return_value.put.return_value = mock_response

                result = client.upload_file("drive_1", "test.pdf", "Folder/test.pdf")
                assert result["id"] == "file_123"

    def test_upload_file_not_found(self):
        """Test upload with missing local file."""
        client = GraphClient("token")

        with pytest.raises(ValueError, match="Failed to read file"):
            client.upload_file("drive_1", "/nonexistent/file.pdf", "target.pdf")

    def test_upload_creates_folder_path(self):
        """Test upload creates nested folder structure."""
        client = GraphClient("token")

        with patch('builtins.open', create=True) as mock_open:
            mock_open.return_value.__enter__.return_value.read.return_value = b"content"

            with patch.object(client, '_ensure_folder_path') as mock_ensure:
                mock_ensure.return_value = "folder_parent_id"

                with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
                    mock_response = MagicMock()
                    mock_response.status_code = 201
                    mock_response.json.return_value = {"id": "file_123"}
                    mock_client.return_value.__enter__.return_value.put.return_value = mock_response

                    client.upload_file("drive_1", "test.pdf", "New/Folder/test.pdf")
                    mock_ensure.assert_called_once_with("drive_1", "New/Folder")


class TestGraphClientDownload:
    """Test file download."""

    def test_download_success(self):
        """Test successful file download."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.content = b"file content"
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            content = client.download_file("drive_1", "file_123")
            assert content == b"file content"

            call_args = mock_client.return_value.__enter__.return_value.get.call_args
            assert call_args[1]["follow_redirects"] is True


class TestGraphClientDelete:
    """Test item deletion."""

    def test_delete_success(self):
        """Test successful delete."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 204
            mock_client.return_value.__enter__.return_value.delete.return_value = mock_response

            # Should not raise
            client.delete_item("drive_1", "file_123")

    def test_delete_missing_file_no_error(self):
        """Test deleting missing file doesn't error (404 is ok)."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 404
            mock_client.return_value.__enter__.return_value.delete.return_value = mock_response

            # Should not raise
            client.delete_item("drive_1", "missing_file")


class TestGraphClientErrorHandling:
    """Test error handling for all HTTP status codes."""

    def test_401_unauthorized(self):
        """Test 401 Unauthorized error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 401
            mock_response.raise_for_status.side_effect = httpx.HTTPError("401")
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Unauthorized|token may be expired"):
                client._get("/me")

    def test_403_forbidden(self):
        """Test 403 Forbidden error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 403
            mock_response.raise_for_status.side_effect = httpx.HTTPError("403")
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Forbidden"):
                client._get("/protected")

    def test_404_not_found(self):
        """Test 404 Not Found error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 404
            mock_response.raise_for_status.side_effect = httpx.HTTPError("404")
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Not found"):
                client._get("/missing")

    def test_429_rate_limit(self):
        """Test 429 Rate Limit error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 429
            mock_response.raise_for_status.side_effect = httpx.HTTPError("429")
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Rate limited"):
                client._get("/items")

    def test_500_server_error(self):
        """Test 500 Server Error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 500
            mock_response.raise_for_status.side_effect = httpx.HTTPError("500")
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Graph API error"):
                client._get("/items")


class TestGraphClientListItems:
    """Test listing folder items."""

    def test_list_items_root(self):
        """Test listing drive root."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {
                "value": [{"id": "item_1", "name": "file.pdf"}]
            }
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            items = client.list_items("drive_1")
            assert len(items) == 1

    def test_list_items_subfolder(self):
        """Test listing specific folder."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"value": []}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            items = client.list_items("drive_1", "folder_123")
            assert len(items) == 0

    def test_list_items_empty_folder(self):
        """Test empty folder returns no items."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"value": []}
            mock_client.return_value.__enter__.return_value.get.return_value = mock_response

            items = client.list_items("drive_1")
            assert items == []


class TestGraphClientCreateFolder:
    """Test folder creation."""

    def test_create_folder_success(self):
        """Test successful folder creation."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 201
            mock_response.json.return_value = {"id": "folder_123", "name": "NewFolder"}
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            result = client.create_folder("drive_1", "root", "NewFolder")
            assert result["id"] == "folder_123"

    def test_create_folder_existing(self):
        """Test creating folder with conflict behavior (rename)."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 201
            mock_response.json.return_value = {"id": "folder_456", "name": "NewFolder1"}
            mock_client.return_value.__enter__.return_value.post.return_value = mock_response

            result = client.create_folder("drive_1", "root", "NewFolder")
            # Should use conflictBehavior: rename
            call_args = mock_client.return_value.__enter__.return_value.post.call_args
            body = call_args[1]["json"]
            assert body["@microsoft.graph.conflictBehavior"] == "rename"


class TestGraphClientMoveItem:
    """Test moving items."""

    def test_move_item_success(self):
        """Test successful move."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"id": "item_1", "name": "file.pdf"}
            mock_client.return_value.__enter__.return_value.patch.return_value = mock_response

            result = client.move_item("drive_1", "item_1", "folder_2")
            assert result["id"] == "item_1"

    def test_move_and_rename(self):
        """Test move with rename."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"id": "item_1", "name": "renamed.pdf"}
            mock_client.return_value.__enter__.return_value.patch.return_value = mock_response

            client.move_item("drive_1", "item_1", "folder_2", "renamed.pdf")

            call_args = mock_client.return_value.__enter__.return_value.patch.call_args
            body = call_args[1]["json"]
            assert body["name"] == "renamed.pdf"
