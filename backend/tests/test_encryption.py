"""Tests for encryption."""
import pytest
from cryptography.fernet import Fernet

from encryption import TokenEncryption


@pytest.fixture
def valid_key():
    """Generate valid Fernet key."""
    return Fernet.generate_key().decode()


@pytest.fixture
def encryption(valid_key):
    """Create encryption instance."""
    return TokenEncryption(valid_key)


def test_encrypt_token(encryption):
    """encrypt returns encrypted token."""
    token = "secret_access_token_123"
    encrypted = encryption.encrypt(token)
    
    # Encrypted should be different from plaintext
    assert encrypted != token
    # Encrypted should be base64 (alphanumeric + - _)
    assert all(c.isalnum() or c in "-_" for c in encrypted)


def test_decrypt_token(encryption):
    """decrypt recovers original token."""
    token = "secret_access_token_123"
    encrypted = encryption.encrypt(token)
    decrypted = encryption.decrypt(encrypted)
    
    assert decrypted == token


def test_decrypt_different_token_fails(encryption):
    """decrypt with different key fails."""
    token = "secret_token"
    encrypted = encryption.encrypt(token)
    
    # Create different encryption instance
    other_key = Fernet.generate_key().decode()
    other_encryption = TokenEncryption(other_key)
    
    with pytest.raises(ValueError, match="Failed to decrypt token"):
        other_encryption.decrypt(encrypted)


def test_encrypt_empty_string(encryption):
    """encrypt handles empty string."""
    encrypted = encryption.encrypt("")
    decrypted = encryption.decrypt(encrypted)
    
    assert decrypted == ""


def test_encrypt_unicode_token(encryption):
    """encrypt handles unicode characters."""
    token = "token_with_unicode_🔐"
    encrypted = encryption.encrypt(token)
    decrypted = encryption.decrypt(encrypted)
    
    assert decrypted == token


def test_invalid_key_raises_error():
    """Invalid encryption key raises error."""
    with pytest.raises(ValueError, match="Invalid TOKEN_ENCRYPTION_KEY"):
        TokenEncryption("not_valid_base64")


def test_decrypt_malformed_encrypted_raises_error(encryption):
    """Decrypt malformed encrypted text raises error."""
    with pytest.raises(ValueError, match="Failed to decrypt token"):
        encryption.decrypt("not_valid_encrypted_data")
