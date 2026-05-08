"""Tests for OAuth flow and authorization."""
import pytest
from unittest.mock import MagicMock, patch
from datetime import datetime, timedelta


class TestMicrosoftOAuthManager:
    """Test OAuth authorization flow."""

    def test_generate_authorization_url_has_required_params(self, mock_config, token_encryption, test_repository):
        """Test OAuth authorization URL includes correct parameters."""
        from oauth import MicrosoftOAuthManager
        
        with patch('oauth.get_config', return_value=mock_config):
            with patch('oauth.get_repository', return_value=test_repository):
                manager = MicrosoftOAuthManager(token_encryption)
                auth_url, state = manager.generate_authorization_url("user_123")
                
                assert "client_id=test_client_id" in auth_url
                assert "response_type=code" in auth_url
                assert "offline_access" in auth_url
                assert "User.Read" in auth_url

    def test_authorization_url_stores_state(self, mock_config, token_encryption, test_repository):
        """Test authorization URL generation stores state server-side."""
        from oauth import MicrosoftOAuthManager
        
        with patch('oauth.get_config', return_value=mock_config):
            with patch('oauth.get_repository', return_value=test_repository):
                manager = MicrosoftOAuthManager(token_encryption)
                auth_url, state = manager.generate_authorization_url("user_123")
                
                stored_state = test_repository.get_oauth_state(state)
                assert stored_state is not None
                assert stored_state["user_id"] == "user_123"

    def test_callback_rejects_invalid_state(self, mock_config, token_encryption, test_repository):
        """Test callback rejects invalid state parameter."""
        from oauth import MicrosoftOAuthManager
        
        with patch('oauth.get_config', return_value=mock_config):
            with patch('oauth.get_repository', return_value=test_repository):
                manager = MicrosoftOAuthManager(token_encryption)
                
                with pytest.raises(ValueError, match="Invalid or expired state"):
                    manager.handle_callback("code_123", "invalid_state", "user_123")

    def test_callback_stores_encrypted_tokens(self, mock_config, token_encryption, test_repository):
        """Test callback stores tokens securely encrypted."""
        from oauth import MicrosoftOAuthManager
        
        with patch('oauth.get_config', return_value=mock_config):
            with patch('oauth.get_repository', return_value=test_repository):
                manager = MicrosoftOAuthManager(token_encryption)
                auth_url, state = manager.generate_authorization_url("user_123")
                
                with patch.object(manager, '_exchange_code_for_tokens') as mock_exchange:
                    with patch.object(manager, '_get_user_info') as mock_userinfo:
                        mock_exchange.return_value = {
                            "access_token": "new_access_token",
                            "refresh_token": "new_refresh_token",
                            "expires_in": 3600,
                        }
                        mock_userinfo.return_value = {"id": "ms_user_123"}
                        
                        connection = manager.handle_callback("code_123", state, "user_123")
                        
                        # Verify encrypted tokens
                        decrypted_access = token_encryption.decrypt(
                            connection.access_token_encrypted
                        )
                        assert decrypted_access == "new_access_token"
