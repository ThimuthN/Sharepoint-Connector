"""Tests for encrypted token storage."""
import pytest
import json
import tempfile
from pathlib import Path
from rpa_sharepoint_connector.token_store import TokenStore
from cryptography.fernet import Fernet


class TestTokenStoreInitialization:
    """Test token store initialization."""

    def test_init_creates_store_dir(self):
        """Test initialization creates store directory."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            assert Path(tmpdir).exists()

    def test_init_generates_encryption_key(self):
        """Test initialization generates encryption key."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            key_file = Path(tmpdir) / ".key"
            assert key_file.exists()

    def test_init_with_provided_key(self):
        """Test initialization with provided key."""
        key = Fernet.generate_key().decode()
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir, encryption_key=key)
            assert store.encryption_key == key

    def test_init_with_invalid_key_raises_error(self):
        """Test invalid key raises clear error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            with pytest.raises(ValueError, match="Invalid encryption key"):
                TokenStore(store_dir=tmpdir, encryption_key="invalid")

    def test_init_loads_existing_key(self):
        """Test reusing existing key."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create first store
            store1 = TokenStore(store_dir=tmpdir)
            key1 = store1.encryption_key

            # Create second store with same dir
            store2 = TokenStore(store_dir=tmpdir)
            key2 = store2.encryption_key

            assert key1 == key2


class TestProfileSaveAndLoad:
    """Test saving and loading profiles."""

    def test_save_profile_encrypts_tokens(self):
        """Test that tokens are encrypted when saved."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            profile = {
                "access_token": "secret_token_123",
                "refresh_token": "secret_refresh_456",
                "user_id": "user_123"
            }

            store.save_profile("test", profile)

            # Check file exists and tokens are encrypted
            profile_file = Path(tmpdir) / "test.json"
            assert profile_file.exists()

            content = json.loads(profile_file.read_text())
            assert content["access_token"] != "secret_token_123"
            assert content["refresh_token"] != "secret_refresh_456"

    def test_load_profile_decrypts_tokens(self):
        """Test that tokens are decrypted when loaded."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            original = {
                "access_token": "secret_token_123",
                "refresh_token": "secret_refresh_456",
                "user_id": "user_123"
            }

            store.save_profile("test", original)
            loaded = store.load_profile("test")

            assert loaded["access_token"] == "secret_token_123"
            assert loaded["refresh_token"] == "secret_refresh_456"
            assert loaded["user_id"] == "user_123"

    def test_load_missing_profile_returns_none(self):
        """Test loading non-existent profile returns None."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            result = store.load_profile("nonexistent")
            assert result is None

    def test_load_corrupted_profile_raises_error(self):
        """Test corrupted profile file raises clear error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)

            # Write corrupted encrypted data
            profile_file = Path(tmpdir) / "corrupted.json"
            profile_file.write_text('{"access_token": "invalid_base64_data"}')

            with pytest.raises(ValueError, match="Failed to load profile"):
                store.load_profile("corrupted")

    def test_load_invalid_json_raises_error(self):
        """Test invalid JSON in profile raises error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)

            profile_file = Path(tmpdir) / "broken.json"
            profile_file.write_text("not valid json {")

            with pytest.raises(ValueError, match="Failed to load profile"):
                store.load_profile("broken")


class TestProfileDeletion:
    """Test profile deletion."""

    def test_delete_profile(self):
        """Test deleting a profile."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            store.save_profile("test", {"data": "value"})

            profile_file = Path(tmpdir) / "test.json"
            assert profile_file.exists()

            store.delete_profile("test")
            assert not profile_file.exists()

    def test_delete_missing_profile_no_error(self):
        """Test deleting non-existent profile doesn't error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            # Should not raise
            store.delete_profile("nonexistent")


class TestProfileListing:
    """Test listing profiles."""

    def test_list_profiles(self):
        """Test listing multiple profiles."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            store.save_profile("profile1", {"data": "1"})
            store.save_profile("profile2", {"data": "2"})
            store.save_profile("profile3", {"data": "3"})

            profiles = store.list_profiles()
            assert len(profiles) == 3
            assert "profile1" in profiles
            assert "profile2" in profiles
            assert "profile3" in profiles

    def test_list_profiles_empty(self):
        """Test listing when no profiles exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            profiles = store.list_profiles()
            assert len(profiles) == 0

    def test_list_profiles_excludes_hidden_files(self):
        """Test that .key and other hidden files are excluded."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            store.save_profile("test", {"data": "value"})

            profiles = store.list_profiles()
            assert ".key" not in profiles


class TestEncryption:
    """Test encryption/decryption."""

    def test_encrypt_decrypt_roundtrip(self):
        """Test encrypting and decrypting text."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)

            original = "This is a secret token value"
            encrypted = store._encrypt(original)
            decrypted = store._decrypt(encrypted)

            assert encrypted != original
            assert decrypted == original

    def test_decrypt_invalid_data_raises_error(self):
        """Test decrypting invalid data raises error."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)

            with pytest.raises(ValueError, match="Failed to decrypt"):
                store._decrypt("invalid_base64_data")

    def test_different_keys_cant_decrypt(self):
        """Test that different encryption keys can't decrypt."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create and encrypt with key1
            store1 = TokenStore(store_dir=tmpdir)
            encrypted = store1._encrypt("secret")

            # Try to decrypt with key2
            key2 = Fernet.generate_key().decode()
            store2 = TokenStore(store_dir=tmpdir, encryption_key=key2)

            with pytest.raises(ValueError):
                store2._decrypt(encrypted)


class TestProfilePermissions:
    """Test file permissions for security."""

    def test_key_file_restrictive_permissions(self):
        """Test key file has restrictive permissions (600)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            key_file = Path(tmpdir) / ".key"

            # On Unix: check permissions are 600 (owner read+write only)
            # Note: chmod may not work on Windows, so this test skips on Windows
            import sys
            import os
            if sys.platform != "win32" and hasattr(os, 'chmod'):
                stat_info = key_file.stat()
                mode = stat_info.st_mode & 0o777
                assert mode == 0o600, f"Key file has wrong permissions: {oct(mode)}"
            else:
                # On Windows, just verify file exists
                assert key_file.exists()

    def test_profile_file_restrictive_permissions(self):
        """Test profile file has restrictive permissions (600)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            store = TokenStore(store_dir=tmpdir)
            store.save_profile("test", {"data": "value"})

            profile_file = Path(tmpdir) / "test.json"
            # On Unix: check permissions are 600 (owner read+write only)
            # Note: chmod may not work on Windows
            import sys
            import os
            if sys.platform != "win32" and hasattr(os, 'chmod'):
                stat_info = profile_file.stat()
                mode = stat_info.st_mode & 0o777
                assert mode == 0o600, f"Profile file has wrong permissions: {oct(mode)}"
            else:
                # On Windows, just verify file exists
                assert profile_file.exists()
