"""Tests for API routes."""
import pytest
from fastapi.testclient import TestClient
from unittest.mock import MagicMock, patch
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from app import app


client = TestClient(app)


class TestMicrosoftConnectorRoutes:
    """Test API endpoints."""

    def test_health_check(self):
        """Test health check endpoint."""
        response = client.get("/health")
        assert response.status_code == 200
        assert response.json()["status"] == "ok"

    def test_connect_endpoint(self):
        """Test /connect endpoint returns auth URL."""
        with patch('routes.oauth_manager') as mock_oauth:
            mock_oauth.generate_authorization_url.return_value = (
                "https://login.microsoftonline.com/...",
                "state_123"
            )
            
            response = client.get("/api/integrations/microsoft/connect")
            assert response.status_code == 200
            data = response.json()
            assert "auth_url" in data
            assert "state" in data

    def test_status_disconnected(self):
        """Test status endpoint when not connected."""
        with patch('routes.repo') as mock_repo:
            mock_repo.get_by_user_id.return_value = None
            
            response = client.get("/api/integrations/microsoft/status")
            assert response.status_code == 200
            data = response.json()
            assert data["connected"] is False

    def test_status_connected(self):
        """Test status endpoint when connected."""
        from models import MicrosoftConnection
        from datetime import datetime, timedelta
        from cryptography.fernet import Fernet
        
        key = Fernet.generate_key().decode()
        from encryption import TokenEncryption
        encryption = TokenEncryption(key)
        
        connection = MicrosoftConnection(
            id="conn_123",
            user_id="demo_user",
            tenant_id="tenant_456",
            microsoft_user_id="ms_user_789",
            access_token_encrypted=encryption.encrypt("token"),
            refresh_token_encrypted=encryption.encrypt("refresh"),
            expires_at=datetime.utcnow() + timedelta(hours=1),
            scopes="offline_access User.Read Sites.Read.All Files.Read.All",
            is_connected=True,
        )
        
        with patch('routes.repo') as mock_repo:
            mock_repo.get_by_user_id.return_value = connection
            
            response = client.get("/api/integrations/microsoft/status")
            assert response.status_code == 200
            data = response.json()
            assert data["connected"] is True
            assert data["microsoft_user_id"] == "ms_user_789"

    def test_sites_requires_connection(self):
        """Test sites endpoint returns 401 when not connected."""
        with patch('routes.token_provider') as mock_provider:
            mock_provider.get_valid_access_token.side_effect = ValueError("No connection")
            
            response = client.get("/api/integrations/microsoft/sites")
            assert response.status_code == 401

    def test_sites_returns_list(self):
        """Test sites endpoint returns site list."""
        with patch('routes.token_provider') as mock_provider:
            with patch('routes.MicrosoftGraphClient') as mock_client_class:
                mock_provider.get_valid_access_token.return_value = "token_123"
                
                mock_client = MagicMock()
                mock_client.list_sites.return_value = [
                    {"id": "site_1", "name": "Team Site"},
                    {"id": "site_2", "name": "Project Site"},
                ]
                mock_client_class.return_value = mock_client
                
                response = client.get("/api/integrations/microsoft/sites")
                assert response.status_code == 200
                data = response.json()
                assert len(data["value"]) == 2

    def test_drives_endpoint(self):
        """Test drives endpoint for a site."""
        with patch('routes.token_provider') as mock_provider:
            with patch('routes.MicrosoftGraphClient') as mock_client_class:
                mock_provider.get_valid_access_token.return_value = "token_123"
                
                mock_client = MagicMock()
                mock_client.list_drives.return_value = [
                    {"id": "drive_1", "name": "Shared Documents"},
                ]
                mock_client_class.return_value = mock_client
                
                response = client.get("/api/integrations/microsoft/sites/site_1/drives")
                assert response.status_code == 200
                mock_client.list_drives.assert_called_once_with("site_1")

    def test_items_endpoint(self):
        """Test items endpoint for a drive."""
        with patch('routes.token_provider') as mock_provider:
            with patch('routes.MicrosoftGraphClient') as mock_client_class:
                mock_provider.get_valid_access_token.return_value = "token_123"
                
                mock_client = MagicMock()
                mock_client.list_drive_items.return_value = [
                    {"id": "item_1", "name": "Document.docx", "file": {}},
                ]
                mock_client_class.return_value = mock_client
                
                response = client.get("/api/integrations/microsoft/drives/drive_1/items")
                assert response.status_code == 200
                mock_client.list_drive_items.assert_called_once_with("drive_1", item_id=None)

    def test_download_endpoint(self):
        """Test download endpoint returns file."""
        with patch('routes.token_provider') as mock_provider:
            with patch('routes.MicrosoftGraphClient') as mock_client_class:
                mock_provider.get_valid_access_token.return_value = "token_123"
                
                mock_client = MagicMock()
                mock_client._get.return_value = {"name": "test.pdf"}
                mock_client.download_drive_item.return_value = b"PDF content"
                mock_client_class.return_value = mock_client
                
                response = client.get("/api/integrations/microsoft/drives/drive_1/items/item_1/download")
                assert response.status_code == 200
                assert "test.pdf" in response.headers.get("content-disposition", "")
