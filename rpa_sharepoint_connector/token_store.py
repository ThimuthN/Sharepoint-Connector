"""Encrypted profile and token storage."""
import json
import logging
import os
from pathlib import Path
from typing import Dict, Optional
from cryptography.fernet import Fernet
import base64

logger = logging.getLogger(__name__)


class TokenStore:
    """Store and retrieve encrypted profiles."""

    INVALID_PROFILE_NAME_CHARS = {"/", "\\", ":"}

    def __init__(self, store_dir: Optional[str] = None, encryption_key: Optional[str] = None):
        """Initialize token store.

        Args:
            store_dir: Directory to store profiles (defaults to ~/.rpa_sharepoint_connector)
            encryption_key: Base64-encoded Fernet key (generates if not provided)
        """
        if store_dir is None:
            store_dir = os.path.expanduser("~/.rpa_sharepoint_connector")
        self.store_dir = Path(store_dir)
        self.store_dir.mkdir(parents=True, exist_ok=True)

        # Load or generate encryption key
        self.key_file = self.store_dir / ".key"
        if encryption_key:
            self.encryption_key = encryption_key
        else:
            self.encryption_key = self._load_or_generate_key()

        try:
            self.cipher = Fernet(self.encryption_key.encode())
        except Exception as e:
            raise ValueError(
                f"Invalid encryption key. Generate with: "
                f"python -c \"from cryptography.fernet import Fernet; "
                f"print(Fernet.generate_key().decode())\""
            ) from e

    def _load_or_generate_key(self) -> str:
        """Load existing key or generate new one."""
        if self.key_file.exists():
            return self.key_file.read_text(encoding="utf-8").strip()

        # Generate new key
        key = Fernet.generate_key().decode()
        self.key_file.write_text(key, encoding="utf-8")
        self.key_file.chmod(0o600)  # Read/write for owner only
        logger.info(f"Generated encryption key at {self.key_file}")
        return key

    def _profile_path(self, profile_name: str) -> Path:
        """Resolve a safe on-disk path for a profile."""
        normalized = (profile_name or "").strip()
        if not normalized:
            raise ValueError("Profile name is required.")
        if normalized in {".", ".."}:
            raise ValueError("Invalid profile name.")
        if normalized.startswith("."):
            raise ValueError("Profile names cannot start with '.'.")
        if any(char in normalized for char in self.INVALID_PROFILE_NAME_CHARS):
            raise ValueError("Profile name contains invalid path characters.")

        return self.store_dir / f"{normalized}.json"

    def save_profile(self, profile_name: str, profile_data: Dict) -> None:
        """Save encrypted profile.

        Args:
            profile_name: Profile identifier
            profile_data: Profile dict with tokens and config
        """
        # Encrypt sensitive fields
        encrypted_data = profile_data.copy()
        if "access_token" in encrypted_data:
            encrypted_data["access_token"] = self._encrypt(encrypted_data["access_token"])
        if "refresh_token" in encrypted_data:
            encrypted_data["refresh_token"] = self._encrypt(encrypted_data["refresh_token"])

        profile_file = self._profile_path(profile_name)
        profile_file.write_text(
            json.dumps(encrypted_data, indent=2, default=str),
            encoding="utf-8",
        )
        profile_file.chmod(0o600)
        logger.info(f"Saved profile: {profile_name}")

    def load_profile(self, profile_name: str) -> Optional[Dict]:
        """Load and decrypt profile.

        Args:
            profile_name: Profile identifier

        Returns:
            Decrypted profile dict or None if not found
        """
        profile_file = self._profile_path(profile_name)
        if not profile_file.exists():
            return None

        try:
            encrypted_data = json.loads(profile_file.read_text(encoding="utf-8"))

            # Decrypt sensitive fields
            if "access_token" in encrypted_data:
                encrypted_data["access_token"] = self._decrypt(encrypted_data["access_token"])
            if "refresh_token" in encrypted_data:
                encrypted_data["refresh_token"] = self._decrypt(encrypted_data["refresh_token"])

            return encrypted_data
        except Exception as e:
            logger.error(f"Failed to load profile {profile_name}: {e}")
            raise ValueError(f"Failed to load profile: {str(e)}")

    def delete_profile(self, profile_name: str) -> None:
        """Delete a profile.

        Args:
            profile_name: Profile identifier
        """
        profile_file = self._profile_path(profile_name)
        if profile_file.exists():
            profile_file.unlink()
            logger.info(f"Deleted profile: {profile_name}")

    def list_profiles(self) -> list:
        """List all saved profiles.

        Returns:
            List of profile names
        """
        return [
            f.stem for f in self.store_dir.glob("*.json")
            if not f.stem.startswith(".")
        ]

    def _encrypt(self, text: str) -> str:
        """Encrypt text."""
        encrypted = self.cipher.encrypt(text.encode())
        return base64.b64encode(encrypted).decode()

    def _decrypt(self, encrypted_text: str) -> str:
        """Decrypt text."""
        try:
            encrypted = base64.b64decode(encrypted_text.encode())
            decrypted = self.cipher.decrypt(encrypted)
            return decrypted.decode()
        except Exception as e:
            raise ValueError(f"Failed to decrypt token: {str(e)}")
