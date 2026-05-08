"""Tests for repositories."""
import pytest
from datetime import datetime, timedelta
from models import MicrosoftConnection
from repositories import InMemoryMicrosoftConnectionRepository


@pytest.fixture
def repo():
    """Create repository."""
    return InMemoryMicrosoftConnectionRepository()


def test_save_connection(repo):
    """save stores connection."""
    connection = MicrosoftConnection(
        id="conn123",
        user_id="user123",
        tenant_id="tenant123",
        microsoft_user_id="ms_user123",
        access_token_encrypted="encrypted_access_token",
        refresh_token_encrypted="encrypted_refresh_token",
        expires_at=datetime.utcnow() + timedelta(hours=1),
        scopes="User.Read Sites.Read.All",
    )
    
    result = repo.save(connection)
    
    assert result.id == "conn123"
    assert repo.connections["conn123"] is not None


def test_get_by_user_id_returns_connection(repo):
    """get_by_user_id retrieves connection."""
    connection = MicrosoftConnection(
        id="conn123",
        user_id="user123",
        tenant_id="tenant123",
        microsoft_user_id="ms_user123",
        access_token_encrypted="encrypted_access_token",
        refresh_token_encrypted="encrypted_refresh_token",
        expires_at=datetime.utcnow() + timedelta(hours=1),
        scopes="User.Read Sites.Read.All",
    )
    repo.save(connection)
    
    retrieved = repo.get_by_user_id("user123")
    
    assert retrieved is not None
    assert retrieved.user_id == "user123"
    assert retrieved.microsoft_user_id == "ms_user123"


def test_get_by_user_id_returns_none_for_missing_user(repo):
    """get_by_user_id returns None for non-existent user."""
    result = repo.get_by_user_id("nonexistent_user")
    
    assert result is None


def test_update_connection(repo):
    """update modifies connection."""
    connection = MicrosoftConnection(
        id="conn123",
        user_id="user123",
        tenant_id="tenant123",
        microsoft_user_id="ms_user123",
        access_token_encrypted="old_token",
        refresh_token_encrypted="old_refresh",
        expires_at=datetime.utcnow() + timedelta(hours=1),
        scopes="User.Read",
        is_connected=True,
    )
    repo.save(connection)
    
    # Modify and update
    connection.access_token_encrypted = "new_token"
    connection.is_connected = False
    repo.update(connection)
    
    retrieved = repo.get_by_user_id("user123")
    assert retrieved.access_token_encrypted == "new_token"
    assert retrieved.is_connected is False


def test_save_oauth_state(repo):
    """save_oauth_state stores state."""
    repo.save_oauth_state("state_123", "user123")
    
    state_data = repo.get_oauth_state("state_123")
    assert state_data is not None
    assert state_data["user_id"] == "user123"


def test_get_oauth_state_returns_none_for_missing_state(repo):
    """get_oauth_state returns None for non-existent state."""
    result = repo.get_oauth_state("nonexistent_state")
    
    assert result is None


def test_delete_oauth_state(repo):
    """delete_oauth_state removes state."""
    repo.save_oauth_state("state_123", "user123")
    repo.delete_oauth_state("state_123")
    
    result = repo.get_oauth_state("state_123")
    assert result is None


def test_delete_nonexistent_oauth_state_is_safe(repo):
    """delete_oauth_state doesn't error on missing state."""
    # Should not raise
    repo.delete_oauth_state("nonexistent_state")
