"""Manage saved SharePoint site/drive/folder configurations."""
import logging
from datetime import datetime
from typing import Dict, Optional
from dataclasses import dataclass, asdict
import json

logger = logging.getLogger(__name__)


@dataclass
class SharePointConfig:
    """A saved SharePoint configuration."""
    site_id: str
    site_name: str
    drive_id: str
    drive_name: str
    folder_id: str
    folder_path: str


class ProfileManager:
    """Manage saved profiles with SharePoint config."""

    def __init__(self, token_store):
        self.token_store = token_store

    def save_profile(
        self,
        profile_name: str,
        access_token: str,
        refresh_token: str,
        expires_at: str,
        site_id: str,
        site_name: str,
        drive_id: str,
        drive_name: str,
        folder_id: str,
        folder_path: str,
        user_id: str,
        user_email: str,
    ) -> None:
        """Save a complete profile with tokens and SharePoint config.

        Args:
            profile_name: Profile identifier
            access_token: Microsoft access token
            refresh_token: Microsoft refresh token
            expires_at: Token expiration timestamp
            site_id: SharePoint site ID
            site_name: SharePoint site name
            drive_id: Document library drive ID
            drive_name: Document library name
            folder_id: Default folder item ID
            folder_path: Default folder path (e.g., "Invoices/Incoming")
            user_id: Microsoft user ID
            user_email: User email address
        """
        profile_data = {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": expires_at,
            "site_id": site_id,
            "site_name": site_name,
            "drive_id": drive_id,
            "drive_name": drive_name,
            "folder_id": folder_id,
            "folder_path": folder_path,
            "user_id": user_id,
            "user_email": user_email,
            "saved_at": datetime.utcnow().isoformat(),
        }
        self.token_store.save_profile(profile_name, profile_data)
        logger.info(f"Profile '{profile_name}' saved for {user_email}")

    def load_profile(self, profile_name: str) -> Optional[Dict]:
        """Load a profile.

        Args:
            profile_name: Profile identifier

        Returns:
            Profile dict or None if not found
        """
        return self.token_store.load_profile(profile_name)

    def delete_profile(self, profile_name: str) -> None:
        """Delete a profile.

        Args:
            profile_name: Profile identifier
        """
        self.token_store.delete_profile(profile_name)
        logger.info(f"Profile '{profile_name}' deleted")

    def list_profiles(self) -> list:
        """List all saved profiles."""
        return self.token_store.list_profiles()

    def get_sharepoint_config(self, profile_name: str) -> Optional[SharePointConfig]:
        """Get SharePoint config from a profile.

        Args:
            profile_name: Profile identifier

        Returns:
            SharePointConfig or None if profile not found
        """
        profile = self.load_profile(profile_name)
        if not profile:
            return None

        return SharePointConfig(
            site_id=profile.get("site_id", ""),
            site_name=profile.get("site_name", ""),
            drive_id=profile.get("drive_id", ""),
            drive_name=profile.get("drive_name", ""),
            folder_id=profile.get("folder_id", ""),
            folder_path=profile.get("folder_path", ""),
        )
