"""Database models."""
from datetime import datetime
from typing import Optional


class MicrosoftConnection:
    """User's Microsoft/SharePoint connection."""
    
    def __init__(
        self,
        id: str,
        user_id: str,
        tenant_id: str,
        microsoft_user_id: str,
        access_token_encrypted: str,
        refresh_token_encrypted: Optional[str],
        expires_at: datetime,
        scopes: str,
        is_connected: bool = True,
        created_at: datetime = None,
        updated_at: datetime = None,
    ):
        self.id = id
        self.user_id = user_id
        self.tenant_id = tenant_id
        self.microsoft_user_id = microsoft_user_id
        self.access_token_encrypted = access_token_encrypted
        self.refresh_token_encrypted = refresh_token_encrypted
        self.expires_at = expires_at
        self.scopes = scopes  # Space-separated scope list
        self.is_connected = is_connected
        self.created_at = created_at or datetime.utcnow()
        self.updated_at = updated_at or datetime.utcnow()
    
    def to_dict(self) -> dict:
        """Convert to dict (safe for JSON, no secrets)."""
        return {
            "id": self.id,
            "user_id": self.user_id,
            "tenant_id": self.tenant_id,
            "microsoft_user_id": self.microsoft_user_id,
            "expires_at": self.expires_at.isoformat(),
            "is_connected": self.is_connected,
            "scopes": self.scopes,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
        }
