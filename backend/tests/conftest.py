"""Pytest fixtures and configuration for backend tests."""
import pytest
from datetime import datetime, timedelta
from unittest.mock import MagicMock
from cryptography.fernet import Fernet
import sys
from pathlib import Path

# Add backend directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models import MicrosoftConnection
from repositories import InMemoryMicrosoftConnectionRepository
from encryption import TokenEncryption


@pytest.fixture(autouse=True)
def reset_repository():
    """Reset the global repository singleton between tests."""
    import repositories
    repositories._repo = None
    yield
    repositories._repo = None


@pytest.fixture
def encryption_key():
    """Generate a valid encryption key for testing."""
    return Fernet.generate_key().decode()


@pytest.fixture
def token_encryption(encryption_key):
    """Create a TokenEncryption instance for testing."""
    return TokenEncryption(encryption_key)


@pytest.fixture
def test_repository():
    """Create a fresh repository for each test."""
    return InMemoryMicrosoftConnectionRepository()


@pytest.fixture
def sample_connection(token_encryption):
    """Create a sample MicrosoftConnection for testing."""
    access_token = "test_access_token_123"
    refresh_token = "test_refresh_token_456"

    return MicrosoftConnection(
        id="conn_123",
        user_id="user_123",
        tenant_id="tenant_456",
        microsoft_user_id="ms_user_789",
        access_token_encrypted=token_encryption.encrypt(access_token),
        refresh_token_encrypted=token_encryption.encrypt(refresh_token),
        expires_at=datetime.utcnow() + timedelta(hours=1),
        scopes="offline_access User.Read Sites.Read.All Files.Read.All",
        is_connected=True,
        created_at=datetime.utcnow(),
        updated_at=datetime.utcnow(),
    )


@pytest.fixture
def expired_connection(token_encryption):
    """Create a connection with expired token."""
    access_token = "expired_access_token"
    refresh_token = "test_refresh_token"

    return MicrosoftConnection(
        id="conn_expired",
        user_id="user_123",
        tenant_id="tenant_456",
        microsoft_user_id="ms_user_789",
        access_token_encrypted=token_encryption.encrypt(access_token),
        refresh_token_encrypted=token_encryption.encrypt(refresh_token),
        expires_at=datetime.utcnow() - timedelta(hours=1),
        scopes="offline_access User.Read Sites.Read.All Files.Read.All",
        is_connected=True,
        created_at=datetime.utcnow(),
        updated_at=datetime.utcnow(),
    )


@pytest.fixture
def disconnected_connection(token_encryption):
    """Create a disconnected connection."""
    access_token = "test_access_token"
    refresh_token = "test_refresh_token"

    return MicrosoftConnection(
        id="conn_disconnected",
        user_id="user_123",
        tenant_id="tenant_456",
        microsoft_user_id="ms_user_789",
        access_token_encrypted=token_encryption.encrypt(access_token),
        refresh_token_encrypted=token_encryption.encrypt(refresh_token),
        expires_at=datetime.utcnow() + timedelta(hours=1),
        scopes="offline_access User.Read Sites.Read.All Files.Read.All",
        is_connected=False,
        created_at=datetime.utcnow(),
        updated_at=datetime.utcnow(),
    )


@pytest.fixture
def mock_config():
    """Create a mock config for testing."""
    config = MagicMock()
    config.MICROSOFT_CLIENT_ID = "test_client_id"
    config.MICROSOFT_CLIENT_SECRET = "test_client_secret"
    config.MICROSOFT_TENANT_ID = "common"
    config.MICROSOFT_REDIRECT_URI = "http://localhost:8000/api/integrations/microsoft/callback"
    config.MICROSOFT_AUTHORIZE_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    config.MICROSOFT_TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    config.MICROSOFT_GRAPH_API = "https://graph.microsoft.com/v1.0"
    config.OAUTH_SCOPES = [
        "offline_access",
        "User.Read",
        "Sites.Read.All",
        "Files.Read.All",
    ]
    return config
