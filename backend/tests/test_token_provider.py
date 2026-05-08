"""Tests for token provider and refresh logic."""
import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta
from token_provider import TokenProvider


class TestTokenProvider:
    """Test token refresh and validation."""

    def test_get_valid_token_no_connection(self, mock_config, token_encryption, test_repository):
        """Test getting token for user with no connection raises error."""
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                with pytest.raises(ValueError, match="No Microsoft connection"):
                    provider.get_valid_access_token("nonexistent_user")

    def test_get_valid_token_disconnected_connection(self, mock_config, token_encryption, test_repository, disconnected_connection):
        """Test getting token from disconnected connection raises error."""
        test_repository.save(disconnected_connection)
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                with pytest.raises(ValueError, match="connection is disconnected"):
                    provider.get_valid_access_token("user_123")

    def test_get_valid_token_not_expired(self, mock_config, token_encryption, test_repository, sample_connection):
        """Test getting valid token when not expired returns it without refresh."""
        test_repository.save(sample_connection)
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                token = provider.get_valid_access_token("user_123")
                assert token == "test_access_token_123"

    def test_get_valid_token_near_expiry_refreshes(self, mock_config, token_encryption, test_repository):
        """Test getting token near expiry triggers refresh."""
        from models import MicrosoftConnection
        
        near_expiry_connection = MicrosoftConnection(
            id="conn_123",
            user_id="user_123",
            tenant_id="tenant_456",
            microsoft_user_id="ms_user_789",
            access_token_encrypted=token_encryption.encrypt("old_token"),
            refresh_token_encrypted=token_encryption.encrypt("refresh_token"),
            expires_at=datetime.utcnow() + timedelta(minutes=2),
            scopes="offline_access User.Read Sites.Read.All Files.Read.All",
            is_connected=True,
        )
        test_repository.save(near_expiry_connection)
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                with patch.object(provider, '_refresh_token') as mock_refresh:
                    mock_refresh.side_effect = lambda conn: setattr(
                        conn, 'access_token_encrypted',
                        token_encryption.encrypt("refreshed_token")
                    ) or setattr(
                        conn, 'expires_at',
                        datetime.utcnow() + timedelta(hours=1)
                    )
                    
                    token = provider.get_valid_access_token("user_123")
                    mock_refresh.assert_called_once()

    def test_refresh_invalid_grant_marks_disconnected(self, mock_config, token_encryption, test_repository, sample_connection):
        """Test invalid_grant error during refresh marks connection disconnected."""
        test_repository.save(sample_connection)
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                import httpx
                with patch('token_provider.httpx.Client') as mock_client:
                    mock_response = MagicMock()
                    mock_response.json.return_value = {"error": "invalid_grant"}
                    error = httpx.HTTPError("Invalid grant")
                    error.response = mock_response
                    mock_response.raise_for_status.side_effect = error
                    mock_client.return_value.__enter__.return_value.post.return_value = mock_response
                    
                    with pytest.raises(ValueError, match="Refresh token is invalid"):
                        provider._refresh_token(sample_connection)
                    
                    # Check connection marked as disconnected
                    updated = test_repository.get_by_user_id("user_123")
                    assert updated.is_connected is False

    def test_refresh_successful_updates_tokens(self, mock_config, token_encryption, test_repository, sample_connection):
        """Test successful refresh updates encrypted tokens and expiry."""
        test_repository.save(sample_connection)
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=test_repository):
                provider = TokenProvider(token_encryption)
                
                with patch('token_provider.httpx.Client') as mock_client:
                    mock_response = MagicMock()
                    mock_response.json.return_value = {
                        "access_token": "new_access_token",
                        "refresh_token": "new_refresh_token",
                        "expires_in": 3600,
                    }
                    mock_client.return_value.__enter__.return_value.post.return_value = mock_response
                    
                    provider._refresh_token(sample_connection)
                    
                    updated = test_repository.get_by_user_id("user_123")
                    decrypted_access = token_encryption.decrypt(updated.access_token_encrypted)
                    assert decrypted_access == "new_access_token"

    def test_refresh_no_refresh_token_raises_error(self, mock_config, token_encryption):
        """Test refresh fails gracefully when no refresh_token is stored."""
        from models import MicrosoftConnection
        
        no_refresh = MicrosoftConnection(
            id="conn_123",
            user_id="user_123",
            tenant_id="tenant_456",
            microsoft_user_id="ms_user_789",
            access_token_encrypted=token_encryption.encrypt("access"),
            refresh_token_encrypted=None,
            expires_at=datetime.utcnow() - timedelta(hours=1),
            scopes="offline_access User.Read Sites.Read.All Files.Read.All",
        )
        
        with patch('token_provider.get_config', return_value=mock_config):
            with patch('token_provider.get_repository', return_value=MagicMock()):
                provider = TokenProvider(token_encryption)
                
                with pytest.raises(ValueError, match="no refresh token"):
                    provider._refresh_token(no_refresh)
