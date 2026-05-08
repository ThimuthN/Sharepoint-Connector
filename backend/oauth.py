"""OAuth authorization flow for Microsoft Graph."""
import uuid
import secrets
import json
import logging
from datetime import datetime, timedelta
from typing import Dict, Tuple, Optional
import httpx
from urllib.parse import urlencode, parse_qs
from repositories import get_repository
from encryption import TokenEncryption
from models import MicrosoftConnection


logger = logging.getLogger(__name__)


class MicrosoftOAuthManager:
    """Manage OAuth flow with Microsoft identity platform."""
    
    def __init__(self, encryption: TokenEncryption, config=None):
        if config is None:
            from config import get_config
            self.config = get_config()
        else:
            self.config = config
        self.encryption = encryption
        self.repo = get_repository()
    
    def generate_authorization_url(self, user_id: str) -> Tuple[str, str]:
        """Generate Microsoft OAuth authorization URL.
        
        Args:
            user_id: Local user identifier
            
        Returns:
            Tuple of (auth_url, state)
        """
        state = secrets.token_urlsafe(32)
        
        # Store state server-side for validation
        self.repo.save_oauth_state(state, user_id)
        
        params = {
            "client_id": self.config.MICROSOFT_CLIENT_ID,
            "redirect_uri": self.config.MICROSOFT_REDIRECT_URI,
            "response_type": "code",
            "scope": " ".join(self.config.OAUTH_SCOPES),
            "state": state,
            "prompt": "select_account",
        }
        
        auth_url = f"{self.config.MICROSOFT_AUTHORIZE_URL}?{urlencode(params)}"
        return auth_url, state
    
    def handle_callback(
        self,
        code: str,
        state: str,
        user_id: str,
    ) -> MicrosoftConnection:
        """Exchange authorization code for tokens.
        
        Args:
            code: Authorization code from Microsoft
            state: State parameter for validation
            user_id: Local user identifier
            
        Returns:
            MicrosoftConnection with encrypted tokens
            
        Raises:
            ValueError: If state is invalid or token exchange fails
        """
        # Validate state
        state_data = self.repo.get_oauth_state(state)
        if not state_data:
            raise ValueError("Invalid or expired state parameter")
        
        if state_data["user_id"] != user_id:
            raise ValueError("State user_id mismatch")
        
        # Exchange code for tokens
        token_response = self._exchange_code_for_tokens(code)
        
        # Get user info from Microsoft
        access_token = token_response["access_token"]
        user_info = self._get_user_info(access_token)
        
        # Create and store connection
        connection = MicrosoftConnection(
            id=str(uuid.uuid4()),
            user_id=user_id,
            tenant_id=token_response.get("tenant_id", self.config.MICROSOFT_TENANT_ID),
            microsoft_user_id=user_info["id"],
            access_token_encrypted=self.encryption.encrypt(access_token),
            refresh_token_encrypted=self.encryption.encrypt(
                token_response.get("refresh_token", "")
            ),
            expires_at=datetime.utcnow() + timedelta(
                seconds=token_response.get("expires_in", 3600)
            ),
            scopes=" ".join(self.config.OAUTH_SCOPES),
        )
        
        self.repo.save(connection)
        
        # Clean up state
        self.repo.delete_oauth_state(state)
        
        logger.info(
            f"OAuth callback successful for user {user_id}, "
            f"Microsoft user {user_info['id']}"
        )
        
        return connection
    
    def _exchange_code_for_tokens(self, code: str) -> Dict:
        """Exchange authorization code for access and refresh tokens.
        
        Args:
            code: Authorization code from Microsoft
            
        Returns:
            Token response dict
            
        Raises:
            ValueError: If token exchange fails
        """
        data = {
            "client_id": self.config.MICROSOFT_CLIENT_ID,
            "client_secret": self.config.MICROSOFT_CLIENT_SECRET,
            "code": code,
            "redirect_uri": self.config.MICROSOFT_REDIRECT_URI,
            "grant_type": "authorization_code",
            "scope": " ".join(self.config.OAUTH_SCOPES),
        }
        
        try:
            with httpx.Client() as client:
                response = client.post(self.config.MICROSOFT_TOKEN_URL, data=data)
                response.raise_for_status()
                return response.json()
        except Exception as e:
            logger.error(f"Token exchange failed: {e}")
            raise ValueError(f"Failed to exchange code for tokens: {str(e)}")
    
    def _get_user_info(self, access_token: str) -> Dict:
        """Get user info from Microsoft Graph.
        
        Args:
            access_token: Microsoft access token
            
        Returns:
            User info dict with at least 'id' field
            
        Raises:
            ValueError: If user info fetch fails
        """
        try:
            with httpx.Client() as client:
                response = client.get(
                    f"{self.config.MICROSOFT_GRAPH_API}/me",
                    headers={"Authorization": f"Bearer {access_token}"},
                )
                response.raise_for_status()
                return response.json()
        except Exception as e:
            logger.error(f"Failed to get user info: {e}")
            raise ValueError(f"Failed to get user info from Microsoft Graph: {str(e)}")
