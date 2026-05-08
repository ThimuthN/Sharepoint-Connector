"""Tests for Microsoft Graph API client."""
import pytest
from unittest.mock import MagicMock, patch
from graph_client import MicrosoftGraphClient


class TestMicrosoftGraphClient:
    """Test Graph API client methods."""

    def test_get_me(self, mock_config):
        """Test getting current user information."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"id": "user_123"}
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                result = client.get_me()
                assert result["id"] == "user_123"

    def test_list_sites_with_search(self, mock_config):
        """Test listing sites with search query."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"value": [{"id": "site_1"}]}
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                sites = client.list_sites(search="Team")
                assert len(sites) == 1

    def test_list_drives(self, mock_config):
        """Test listing drives for a site."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"value": [{"id": "drive_1"}]}
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                drives = client.list_drives("site_1")
                assert len(drives) == 1

    def test_list_drive_items(self, mock_config):
        """Test listing items in a drive."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"value": [{"id": "item_1"}]}
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                items = client.list_drive_items("drive_1")
                assert len(items) == 1

    def test_download_file(self, mock_config):
        """Test downloading a file."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"id": "file_1"}
                mock_content = MagicMock()
                mock_content.content = b"file content"
                
                mock_client_class.return_value.__enter__.return_value.get.side_effect = [
                    mock_response, mock_content
                ]
                
                content = client.download_drive_item("drive_1", "file_1")
                assert content == b"file content"

    def test_download_folder_error(self, mock_config):
        """Test downloading folder raises error."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token_123")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.json.return_value = {"id": "folder", "folder": {}}
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                with pytest.raises(ValueError, match="Cannot download folder"):
                    client.download_drive_item("drive_1", "folder")

    def test_401_error(self, mock_config):
        """Test 401 unauthorized error."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("bad_token")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.status_code = 401
                mock_response.raise_for_status.side_effect = Exception("401")
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                with pytest.raises(ValueError, match="Unauthorized"):
                    client._get("/me")

    def test_429_rate_limit(self, mock_config):
        """Test 429 rate limit error."""
        with patch('graph_client.get_config', return_value=mock_config):
            client = MicrosoftGraphClient("token")
            
            with patch('graph_client.httpx.Client') as mock_client_class:
                mock_response = MagicMock()
                mock_response.status_code = 429
                mock_response.raise_for_status.side_effect = Exception("429")
                mock_client_class.return_value.__enter__.return_value.get.return_value = mock_response
                
                with pytest.raises(ValueError, match="Rate limited"):
                    client._get("/users")
