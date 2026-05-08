"""Token provider for getting valid access tokens."""
import logging
from datetime import datetime, timedelta
from typing import Optional
import httpx
from repositories import get_repository
from encryption import TokenEncryption
from models import MicrosoftConnection


logger = logging.getLogger(__name__)


class TokenProvider:
    """Provides valid access tokens, refreshing as needed."""
    
    def __init__(self, encryption: TokenEncryption, config=None):
        if config is None:
            from config import get_config
            self.config = get_config()
        else:
            self.config = config
        self.encryption = encryption
        self.repo = get_repository()
    
    def get_valid_access_token(self, user_id: str) -> str:
        """Get a valid access token for the user.
        
        Refreshes token if expired or near expiry (within 5 minutes).
        
        Args:
            user_id: Local user identifier
            
        Returns:
            Valid access token string
            
        Raises:
            ValueError: If user has no connection or token refresh fails
        """
        connection = self.repo.get_by_user_id(user_id)
        if not connection:
            raise ValueError(f"No Microsoft connection found for user {user_id}")
        
        if not connection.is_connected:
            raise ValueError(f"Microsoft connection is disconnected for user {user_id}")
        
        # Check if token is near expiry (within 5 minutes)
        time_until_expiry = connection.expires_at - datetime.utcnow()
        if time_until_expiry.total_seconds() < 300:
            logger.info(f"Token for user {user_id} near expiry, refreshing...")
            self._refresh_token(connection)
            connection = self.repo.get_by_user_id(user_id)
        
        return self.encryption.decrypt(connection.access_token_encrypted)
    
    def _refresh_token(self, connection: MicrosoftConnection) -> None:
        """Refresh connection's access and refresh tokens.
        
        Args:
            connection: MicrosoftConnection to refresh
            
        Raises:
            ValueError: If refresh fails
        """
        if not connection.refresh_token_encrypted:
            raise ValueError(
                f"Cannot refresh token: no refresh token stored for user {connection.user_id}"
            )
        
        refresh_token = self.encryption.decrypt(connection.refresh_token_encrypted)
        
        data = {
            "client_id": self.config.MICROSOFT_CLIENT_ID,
            "client_secret": self.config.MICROSOFT_CLIENT_SECRET,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
            "scope": " ".join(self.config.OAUTH_SCOPES),
        }
        
        try:
            with httpx.Client() as client:
                response = client.post(self.config.MICROSOFT_TOKEN_URL, data=data)
                # Try to get error body first before raising
                error_body = None
                try:
                    error_body = response.json()
                except:
                    pass
                
                # Check for invalid_grant error
                if error_body and error_body.get("error") == "invalid_grant":
                    logger.warning(
                        f"Refresh token invalid for user {connection.user_id}, "
                        "marking connection disconnected"
                    )
                    connection.is_connected = False
                    self.repo.update(connection)
                    raise ValueError(
                        "Refresh token is invalid. Please reconnect Microsoft account."
                    )
                
                # Now raise for any HTTP errors
                response.raise_for_status()
                token_response = response.json()
        except ValueError:
            # Re-raise ValueError as-is
            raise
        except Exception as e:
            logger.error(f"Token refresh failed: {e}")
            raise ValueError(f"Failed to refresh token: {str(e)}")
        
        # Update connection with new tokens
        connection.access_token_encrypted = self.encryption.encrypt(
            token_response["access_token"]
        )
        connection.refresh_token_encrypted = self.encryption.encrypt(
            token_response.get("refresh_token", refresh_token)
        )
        connection.expires_at = datetime.utcnow() + timedelta(
            seconds=token_response.get("expires_in", 3600)
        )
        
        self.repo.update(connection)
        logger.info(f"Token refreshed successfully for user {connection.user_id}")
