"""Repository for MicrosoftConnection storage and retrieval."""
import json
import uuid
from datetime import datetime
from typing import Optional
from pathlib import Path
from models import MicrosoftConnection


class InMemoryMicrosoftConnectionRepository:
    """In-memory storage for MVP (not for production)."""
    
    def __init__(self):
        self.connections: dict[str, dict] = {}
        self.oauth_states: dict[str, dict] = {}
    
    def save(self, connection: MicrosoftConnection) -> MicrosoftConnection:
        """Save a connection."""
        self.connections[connection.id] = {
            "id": connection.id,
            "user_id": connection.user_id,
            "tenant_id": connection.tenant_id,
            "microsoft_user_id": connection.microsoft_user_id,
            "access_token_encrypted": connection.access_token_encrypted,
            "refresh_token_encrypted": connection.refresh_token_encrypted,
            "expires_at": connection.expires_at.isoformat(),
            "scopes": connection.scopes,
            "is_connected": connection.is_connected,
            "created_at": connection.created_at.isoformat(),
            "updated_at": connection.updated_at.isoformat(),
        }
        return connection
    
    def get_by_user_id(self, user_id: str) -> Optional[MicrosoftConnection]:
        """Get connection by user ID."""
        for conn_data in self.connections.values():
            if conn_data["user_id"] == user_id:
                return self._from_dict(conn_data)
        return None
    
    def update(self, connection: MicrosoftConnection) -> MicrosoftConnection:
        """Update a connection."""
        connection.updated_at = datetime.utcnow()
        return self.save(connection)
    
    def save_oauth_state(self, state: str, user_id: str) -> None:
        """Save OAuth state for validation."""
        self.oauth_states[state] = {
            "user_id": user_id,
            "created_at": datetime.utcnow().isoformat(),
        }
    
    def get_oauth_state(self, state: str) -> Optional[dict]:
        """Get OAuth state."""
        return self.oauth_states.get(state)
    
    def delete_oauth_state(self, state: str) -> None:
        """Delete OAuth state after validation."""
        self.oauth_states.pop(state, None)
    
    @staticmethod
    def _from_dict(data: dict) -> MicrosoftConnection:
        """Convert dict to MicrosoftConnection."""
        return MicrosoftConnection(
            id=data["id"],
            user_id=data["user_id"],
            tenant_id=data["tenant_id"],
            microsoft_user_id=data["microsoft_user_id"],
            access_token_encrypted=data["access_token_encrypted"],
            refresh_token_encrypted=data.get("refresh_token_encrypted"),
            expires_at=datetime.fromisoformat(data["expires_at"]),
            scopes=data["scopes"],
            is_connected=data.get("is_connected", True),
            created_at=datetime.fromisoformat(data["created_at"]),
            updated_at=datetime.fromisoformat(data["updated_at"]),
        )


def get_repository() -> InMemoryMicrosoftConnectionRepository:
    """Get singleton repository instance."""
    global _repo
    if _repo is None:
        _repo = InMemoryMicrosoftConnectionRepository()
    return _repo


_repo: Optional[InMemoryMicrosoftConnectionRepository] = None
