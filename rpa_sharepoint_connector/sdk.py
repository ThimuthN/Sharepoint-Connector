"""Simple SDK for RPA bots to interact with SharePoint."""
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from .auth import MicrosoftAuth
from .token_store import TokenStore
from .graph_client import GraphClient

logger = logging.getLogger(__name__)


class SharePointClient:
    """SharePoint client for bots."""

    def __init__(self, profile: str, store_dir: Optional[str] = None, sharepoint_url: Optional[str] = None):
        self.store = TokenStore(store_dir=store_dir)
        profile_data = self.store.load_profile(profile)
        if not profile_data:
            raise ValueError(
                f"Profile '{profile}' not found. "
                f"Run: python -m rpa_sharepoint_connector configure --profile {profile}"
            )

        self.profile = profile
        self.profile_data = profile_data
        self.sharepoint_url = sharepoint_url

        self.auth = MicrosoftAuth(
            client_id=profile_data.get("client_id"),
            tenant_id=profile_data.get("tenant_id"),
        )
        self._ensure_valid_token()
        self.graph = GraphClient(self.profile_data["access_token"])

        if sharepoint_url:
            self._resolve_drive_from_url()
        else:
            self.drive_id = profile_data["drive_id"]
            self.folder_id = profile_data["folder_id"]

    def close(self) -> None:
        """Close underlying network resources."""
        self.graph.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def _resolve_drive_from_url(self) -> None:
        """Resolve drive_id and folder_id from SharePoint URL."""
        from urllib.parse import urlparse
        try:
            parsed_url = urlparse(self.sharepoint_url)
            hostname = parsed_url.hostname

            if not hostname:
                raise ValueError("Invalid SharePoint URL: missing hostname")

            # Try standard site path first
            try:
                from .cli_setup import parse_sharepoint_url
                parsed = parse_sharepoint_url(self.sharepoint_url)
                site_path = parsed["site_path"]
                library_name = parsed.get("library_name", "Documents")
                folder_path = parsed.get("folder_path", "")
                site = self.graph._get(f"/sites/{hostname}:{site_path}")
            except ValueError:
                # Fall back to root site
                site = self.graph._get(f"/sites/{hostname}:/")
                library_name = "Documents"
                folder_path = ""

            drives = self.graph.list_drives(site["id"])

            from .cli_setup import select_drive
            drive = select_drive(drives, library_name)
            self.drive_id = drive["id"]

            if folder_path:
                self.folder_id = self.graph._ensure_folder_path(drive["id"], folder_path)
            else:
                self.folder_id = "root"

            logger.info(f"Resolved SharePoint URL to drive: {self.drive_id}")
        except Exception as e:
            raise ValueError(f"Failed to resolve SharePoint URL: {e}") from e

    def upload(self, local_path: str, remote_path: str, conflict: str = "overwrite") -> str:
        logger.info(f"Uploading {local_path} to {remote_path} (conflict={conflict})")
        try:
            result = self.graph.upload_file(
                self.drive_id, local_path, remote_path, conflict=conflict
            )
            logger.info(f"Uploaded successfully: {result['id']}")
            return result["id"]
        except Exception as e:
            logger.error(f"Upload failed: {e}")
            raise

    def download(self, remote_path: str, local_path: str) -> None:
        """Download a file from SharePoint."""

    def delete(self, remote_path: str) -> None:
        """Delete a file or folder."""

    def exists(self, remote_path: str) -> bool:
        """Check if file/folder exists."""

    def list(self, folder_path: str = "") -> List[Dict]:
        """List files and folders in a folder."""

    def mkdir(self, folder_path: str) -> str:
        """Create a folder."""

    def move(
        self,
        source_path: str,
        target_path: str,
        new_name: Optional[str] = None,
    ) -> None:
        """Move or rename a file/folder."""

    def _resolve_item_id(self, value: str) -> str:
        """Resolve a user-provided path or ID into a drive item ID.

        Rules:
        - `id:<value>` always means explicit item ID.
        - values containing `/` are treated as paths.
        - single-segment values try path lookup first, then fallback to raw ID.
        """
        candidate = (value or "").strip()
        if not candidate:
            raise ValueError("Path or item ID is required.")

        if candidate.startswith("id:"):
            explicit_id = candidate[3:].strip()
            if not explicit_id:
                raise ValueError("Explicit ID prefix provided but no item ID found.")
            return explicit_id

        if "/" in candidate:
            return self._find_item_id(candidate)

        # Single segment could be a root filename/folder or an item ID.
        try:
            return self._find_item_id(candidate)
        except ValueError:
            return candidate

    def health_check(self) -> Dict[str, bool]:
        """Perform preflight health check.

        Verifies:
        - Profile exists and is valid
        - Token can be refreshed
        - Configured drive and folder exist
        - Basic connectivity to Graph API

        Returns:
            Dict with check results: {"profile": bool, "token": bool, "drive": bool, "folder": bool}

        Raises:
            ValueError: If critical checks fail
        """
        checks = {
            "profile": False,
            "token": False,
            "drive": False,
            "folder": False,
        }

        try:
            # Check profile is loaded
            if not self.profile or not self.profile_data:
                raise ValueError("Profile not loaded")
            checks["profile"] = True
            logger.info(f"Health check: profile '{self.profile}' OK")

            # Check token can be refreshed
            try:
                self._ensure_valid_token()
                checks["token"] = True
                logger.info("Health check: token refresh OK")
            except Exception as e:
                logger.error(f"Health check: token refresh failed: {e}")
                raise ValueError(f"Token refresh failed: {str(e)}")

            # Check drive exists
            try:
                self.graph.get_me()  # Basic connectivity check
                checks["drive"] = True
                logger.info("Health check: drive connectivity OK")
            except Exception as e:
                logger.error(f"Health check: drive connectivity failed: {e}")
                raise ValueError(f"Cannot connect to SharePoint: {str(e)}")

            # Check folder exists
            try:
                if self.folder_id and self.folder_id != "root":
                    self.graph.get_drive_item(self.drive_id, self.folder_id)
                checks["folder"] = True
                logger.info("Health check: folder exists OK")
            except Exception as e:
                logger.warning(f"Health check: folder check failed: {e}")
                # Don't fail on folder check - it might be deleted but drive is still OK
                checks["folder"] = False

            logger.info(f"Health check complete: {checks}")
            return checks

        except Exception as e:
            logger.error(f"Health check failed: {e}")
            raise

    def _find_item_id(self, path: str) -> str:
        """Find item ID from a path."""

    def _build_remote_path(self, path: str) -> str:
        """Build a drive-root-relative path from configured base folder and user input."""
        base_path = (self.profile_data.get("folder_path") or "").strip("/")
        relative_path = (path or "").strip("/")
        parts = [part for part in (base_path, relative_path) if part]
        return "/".join(parts)

    def _ensure_valid_token(self) -> None:
        """Refresh token if expired or near expiry."""
        expires_at = datetime.fromisoformat(self.profile_data["expires_at"])

        if datetime.utcnow() >= expires_at - timedelta(minutes=5):
            logger.info(f"Token expired for {self.profile}, refreshing...")
            try:
                token_response = self.auth.refresh_token(
                    self.profile_data["refresh_token"]
                )
                self.profile_data["access_token"] = token_response["access_token"]
                self.profile_data["refresh_token"] = token_response.get(
                    "refresh_token", self.profile_data["refresh_token"]
                )
                self.profile_data["expires_at"] = (
                    datetime.utcnow() + timedelta(seconds=token_response["expires_in"])
                ).isoformat()

                # Update stored profile
                self.store.save_profile(self.profile, self.profile_data)
                logger.info("Token refreshed successfully")
            except Exception as e:
                logger.error(f"Token refresh failed: {e}")
                raise
