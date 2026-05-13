"""Tests for Microsoft Graph API client."""
import pytest
import tempfile
import os
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
            mock_client.return_value.get.return_value = mock_response

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
            mock_client.return_value.get.return_value = mock_response

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
            mock_client.return_value.get.return_value = mock_response

            drives = client.list_drives("site_123")
            assert len(drives) == 0

    def test_list_drives_follows_pagination(self):
        """Test drive listing follows @odata.nextLink pages."""
        client = GraphClient("token")
        with patch.object(client, "_get") as mock_get:
            with patch.object(client, "_get_absolute") as mock_get_absolute:
                mock_get.return_value = {
                    "value": [{"id": "drive_1"}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/next1",
                }
                mock_get_absolute.side_effect = [
                    {
                        "value": [{"id": "drive_2"}],
                        "@odata.nextLink": "https://graph.microsoft.com/v1.0/next2",
                    },
                    {
                        "value": [{"id": "drive_3"}],
                    },
                ]

                drives = client.list_drives("site_123")
                assert [d["id"] for d in drives] == ["drive_1", "drive_2", "drive_3"]

    def test_list_drives_rejects_invalid_next_link_type(self):
        """Invalid nextLink values should fail fast instead of looping."""
        client = GraphClient("token")
        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {
                "value": [{"id": "drive_1"}],
                "@odata.nextLink": MagicMock(),
            }

            with pytest.raises(ValueError, match="invalid @odata.nextLink"):
                client.list_drives("site_123")

    def test_list_drives_rejects_pagination_loop(self):
        """Repeated nextLink values should fail fast instead of looping forever."""
        client = GraphClient("token")
        with patch.object(client, "_get") as mock_get:
            with patch.object(client, "_get_absolute") as mock_get_absolute:
                mock_get.return_value = {
                    "value": [{"id": "drive_1"}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/next1",
                }
                mock_get_absolute.return_value = {
                    "value": [{"id": "drive_2"}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/next1",
                }

                with pytest.raises(ValueError, match="pagination loop"):
                    client.list_drives("site_123")


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
                mock_client.return_value.put.return_value = mock_response

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
                    mock_client.return_value.put.return_value = mock_response

                    client.upload_file("drive_1", "test.pdf", "New/Folder/test.pdf")
                    mock_ensure.assert_called_once_with("drive_1", "New/Folder")

    def test_upload_large_file_uses_upload_session(self):
        """Large uploads should use createUploadSession + chunk PUT."""
        client = GraphClient("token")

        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            tmp.write(b"ab")
            tmp_path = tmp.name

        try:
            with patch.object(client, "SIMPLE_UPLOAD_LIMIT_BYTES", 1):
                with patch("rpa_sharepoint_connector.graph_client.httpx.Client") as mock_client:
                    session_response = MagicMock()
                    session_response.status_code = 200
                    session_response.json.return_value = {"uploadUrl": "https://upload.example"}

                    chunk_response = MagicMock()
                    chunk_response.status_code = 201
                    chunk_response.json.return_value = {"id": "file_large_1", "name": "big.bin"}

                    http_client = mock_client.return_value
                    http_client.post.return_value = session_response
                    http_client.put.return_value = chunk_response

                    result = client.upload_file("drive_1", tmp_path, "big.bin")
                    assert result["id"] == "file_large_1"

                    post_url = http_client.post.call_args.args[0]
                    assert post_url.endswith("/createUploadSession")
        finally:
            os.unlink(tmp_path)


class TestGraphClientDownload:
    """Test file download."""

    def test_download_success(self):
        """Test successful file download."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.content = b"file content"
            mock_client.return_value.get.return_value = mock_response

            content = client.download_file("drive_1", "file_123")
            assert content == b"file content"

            call_args = mock_client.return_value.get.call_args
            assert call_args[1]["follow_redirects"] is True

    def test_download_to_path_streams_to_temp_file_then_moves(self, tmp_path):
        """Direct-to-path downloads should stream bytes to disk without buffering full content."""
        client = GraphClient("token")
        target_path = tmp_path / "downloaded.bin"

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.iter_bytes.return_value = [b"hello ", b"world"]
            stream_cm = MagicMock()
            stream_cm.__enter__.return_value = mock_response
            stream_cm.__exit__.return_value = False

            http_client = mock_client.return_value
            http_client.stream.return_value = stream_cm

            client.download_file_to_path("drive_1", "file_123", str(target_path))

            assert target_path.read_bytes() == b"hello world"
            call_args = http_client.stream.call_args
            assert call_args.args[0] == "GET"
            assert call_args.args[1].endswith("/drives/drive_1/items/file_123/content")
            assert call_args.kwargs["follow_redirects"] is True


class TestGraphClientDelete:
    """Test item deletion."""

    def test_delete_success(self):
        """Test successful delete."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 204
            mock_client.return_value.delete.return_value = mock_response

            # Should not raise
            client.delete_item("drive_1", "file_123")

    def test_delete_missing_file_no_error(self):
        """Test deleting missing file doesn't error (404 is ok)."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 404
            mock_client.return_value.delete.return_value = mock_response

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
            mock_client.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Unauthorized|token may be expired"):
                client._get("/me")

    def test_403_forbidden(self):
        """Test 403 Forbidden error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 403
            mock_response.raise_for_status.side_effect = httpx.HTTPError("403")
            mock_client.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Forbidden"):
                client._get("/protected")

    def test_404_not_found(self):
        """Test 404 Not Found error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 404
            mock_response.raise_for_status.side_effect = httpx.HTTPError("404")
            mock_client.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Not found"):
                client._get("/missing")

    def test_429_rate_limit(self):
        """Test 429 Rate Limit error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 429
            mock_response.raise_for_status.side_effect = httpx.HTTPError("429")
            mock_client.return_value.get.return_value = mock_response

            with pytest.raises(ValueError, match="Rate limited"):
                client._get("/items")

    def test_500_server_error(self):
        """Test 500 Server Error."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 500
            mock_response.raise_for_status.side_effect = httpx.HTTPError("500")
            mock_client.return_value.get.return_value = mock_response

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
            mock_client.return_value.get.return_value = mock_response

            items = client.list_items("drive_1")
            assert len(items) == 1

    def test_list_items_subfolder(self):
        """Test listing specific folder."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"value": []}
            mock_client.return_value.get.return_value = mock_response

            items = client.list_items("drive_1", "folder_123")
            assert len(items) == 0

    def test_list_items_empty_folder(self):
        """Test empty folder returns no items."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"value": []}
            mock_client.return_value.get.return_value = mock_response

            items = client.list_items("drive_1")
            assert items == []

    def test_list_items_follows_pagination(self):
        """Folder listing should aggregate all nextLink pages."""
        client = GraphClient("token")
        with patch.object(client, "_get") as mock_get:
            with patch.object(client, "_get_absolute") as mock_get_absolute:
                mock_get.return_value = {
                    "value": [{"id": "item_1"}],
                    "@odata.nextLink": "https://graph.microsoft.com/v1.0/next_items",
                }
                mock_get_absolute.return_value = {"value": [{"id": "item_2"}]}

                items = client.list_items("drive_1", "folder_123")
                assert [i["id"] for i in items] == ["item_1", "item_2"]


class TestGraphClientPathResolution:
    """Test direct path-based item resolution."""

    def test_get_item_by_path_uses_root_path_endpoint(self):
        """Path resolution should use a direct Graph path endpoint."""
        client = GraphClient("token")

        with patch.object(client, "_get") as mock_get:
            mock_get.return_value = {"id": "item_123", "name": "Quarterly Report.pdf"}

            result = client.get_item_by_path("drive_1", "Folder A/Quarterly Report.pdf")

            assert result["id"] == "item_123"
            mock_get.assert_called_once_with(
                "/drives/drive_1/root:/Folder%20A/Quarterly%20Report.pdf"
            )

    def test_ensure_folder_path_short_circuits_existing_path(self):
        """Existing folder paths should resolve in one lookup without creating folders."""
        client = GraphClient("token")

        with patch.object(client, "get_item_by_path") as mock_get_item:
            mock_get_item.return_value = {"id": "folder_123", "folder": {}}

            result = client._ensure_folder_path("drive_1", "A/B/C")

            assert result == "folder_123"
            mock_get_item.assert_called_once_with("drive_1", "A/B/C")


class TestGraphClientCreateFolder:
    """Test folder creation."""

    def test_create_folder_success(self):
        """Test successful folder creation."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 201
            mock_response.json.return_value = {"id": "folder_123", "name": "NewFolder"}
            mock_client.return_value.post.return_value = mock_response

            result = client.create_folder("drive_1", "root", "NewFolder")
            assert result["id"] == "folder_123"

    def test_create_folder_existing(self):
        """Test creating folder with conflict behavior (rename)."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 201
            mock_response.json.return_value = {"id": "folder_456", "name": "NewFolder1"}
            mock_client.return_value.post.return_value = mock_response

            result = client.create_folder("drive_1", "root", "NewFolder")
            # Should use conflictBehavior: rename
            call_args = mock_client.return_value.post.call_args
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
            mock_client.return_value.patch.return_value = mock_response

            result = client.move_item("drive_1", "item_1", "folder_2")
            assert result["id"] == "item_1"

    def test_move_and_rename(self):
        """Test move with rename."""
        client = GraphClient("token")

        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            mock_response = MagicMock()
            mock_response.status_code = 200
            mock_response.json.return_value = {"id": "item_1", "name": "renamed.pdf"}
            mock_client.return_value.patch.return_value = mock_response

            client.move_item("drive_1", "item_1", "folder_2", "renamed.pdf")

            call_args = mock_client.return_value.patch.call_args
            body = call_args[1]["json"]
            assert body["name"] == "renamed.pdf"

    def test_reuses_single_http_client_instance(self):
        """Multiple operations should share one lazily-created HTTP client."""
        with patch('rpa_sharepoint_connector.graph_client.httpx.Client') as mock_client:
            http_client = MagicMock()
            mock_client.return_value = http_client

            get_response = MagicMock()
            get_response.status_code = 200
            get_response.json.side_effect = [
                {"id": "user_123"},
                {"value": []},
            ]
            http_client.get.return_value = get_response

            client = GraphClient("token")
            client.get_me()
            client.list_items("drive_1")

            assert mock_client.call_count == 1
            assert http_client.get.call_count == 2
