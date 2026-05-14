"""Simple SDK for RPA bots to interact with SharePoint."""
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from .auth import MicrosoftAuth
from .token_store import TokenStore
from .graph_client import GraphClient

logger = logging.getLogger(__name__)


class SharePointClient:
    """Simple client for bots to interact with SharePoint.

    Usage:
        sp = SharePointClient(profile="client_a")
        sp.upload("local.pdf", "Invoices/Incoming/local.pdf")
        sp.download("Invoices/Incoming/invoice.pdf", "invoice.pdf")
        files = sp.list("Invoices/Incoming")
        sp.delete("Invoices/Temp/old.pdf")
    """

    def __init__(self, profile: str, store_dir: Optional[str] = None, sharepoint_url: Optional[str] = None):
        """Initialize SharePoint client.

        Args:
            profile: Profile name (must be configured first)
            store_dir: Optional token store directory
            sharepoint_url: Optional SharePoint site URL to auto-resolve drive (skips set-target)

        Raises:
            ValueError: If profile not found or tokens invalid
        """
        self.store = TokenStore(store_dir=store_dir)

        # Load profile
        profile_data = self.store.load_profile(profile)
        if not profile_data:
            raise ValueError(
                f"Profile '{profile}' not found. "
                f"Run: python -m rpa_sharepoint_connector configure --profile {profile}"
            )

        self.profile = profile
        self.profile_data = profile_data
        self.sharepoint_url = sharepoint_url

        # Initialize auth using profile-bound app/tenant when available.
        self.auth = MicrosoftAuth(
            client_id=profile_data.get("client_id"),
            tenant_id=profile_data.get("tenant_id"),
        )
        self._ensure_valid_token()
        self.graph = GraphClient(self.profile_data["access_token"])

        # Resolve drive and folder
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

    def upload(
        self,
        local_path: str,
        remote_path: str,
        conflict: str = "overwrite",
    ) -> str:
        """Upload a file to SharePoint with conflict handling.

        Args:
            local_path: Local file path
            remote_path: Remote path (e.g., "Folder/filename.pdf")
            conflict: Conflict behavior:
                - "overwrite": Replace existing file (default for backwards compatibility)
                - "fail_if_exists": Raise error if file exists
                - "rename": Create unique filename if exists (e.g., file (1).pdf)

        Returns:
            Item ID of uploaded file

        Raises:
            ValueError: If upload fails or conflict="fail_if_exists" and file exists
        """
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
        """Download a file from SharePoint.

        Args:
            remote_path: Remote file path or item ID
            local_path: Local file path to save to

        Raises:
            ValueError: If download fails
        """
        logger.info(f"Downloading {remote_path} to {local_path}")
        try:
            item_id = self._resolve_item_id(remote_path)
            self.graph.download_file_to_path(self.drive_id, item_id, local_path)
            logger.info(f"Downloaded successfully to {local_path}")
        except Exception as e:
            logger.error(f"Download failed: {e}")
            raise

    def delete(self, remote_path: str) -> None:
        """Delete a file or folder.

        Args:
            remote_path: Remote file/folder path or item ID

        Raises:
            ValueError: If delete fails or path is dangerous
        """
        # Prevent deletion of root or configured folder
        if not remote_path or remote_path == "/" or remote_path == "root":
            raise ValueError(
                "Cannot delete root folder. Specify the file/folder to delete."
            )

        if remote_path == self.profile_data.get("folder_path", ""):
            raise ValueError(
                f"Cannot delete configured default folder: {remote_path}. "
                "This would break future operations."
            )

        logger.info(f"Deleting {remote_path}")
        try:
            item_id = self._resolve_item_id(remote_path)

            self.graph.delete_item(self.drive_id, item_id)
            logger.info(f"Deleted: {remote_path}")
        except Exception as e:
            logger.error(f"Delete failed: {e}")
            raise

    def exists(self, remote_path: str) -> bool:
        """Check if file/folder exists.

        Args:
            remote_path: Remote path or item ID

        Returns:
            True if exists, False otherwise
        """
        try:
            item_id = self._resolve_item_id(remote_path)
            return self.graph.item_exists(self.drive_id, item_id)
        except ValueError as e:
            if "not found" in str(e).lower():
                return False
            raise

    def list(self, folder_path: str = "") -> List[Dict]:
        """List files and folders in a folder.

        Args:
            folder_path: Folder path (defaults to configured folder)

        Returns:
            List of items with name, id, size, etc.
        """
        try:
            if folder_path:
                parent_id = self._find_item_id(folder_path)
            else:
                parent_id = self.folder_id

            items = self.graph.list_items(self.drive_id, parent_id)
            return [
                {
                    "name": item.get("name"),
                    "id": item.get("id"),
                    "size": item.get("size", 0),
                    "is_folder": "folder" in item,
                }
                for item in items
            ]
        except Exception as e:
            logger.error(f"List failed: {e}")
            raise

    def mkdir(self, folder_path: str) -> str:
        """Create a folder.

        Args:
            folder_path: Path to create (e.g., "Invoices/Temp")

        Returns:
            Item ID of created folder

        Raises:
            ValueError: If creation fails
        """
        logger.info(f"Creating folder: {folder_path}")
        try:
            absolute_path = self._build_remote_path(folder_path)
            parent_id = self.graph._ensure_folder_path(self.drive_id, absolute_path)
            logger.info(f"Created folder: {folder_path}")
            return parent_id
        except Exception as e:
            logger.error(f"Mkdir failed: {e}")
            raise

    def move(
        self,
        source_path: str,
        target_path: str,
        new_name: Optional[str] = None,
    ) -> None:
        """Move or rename a file/folder.

        Args:
            source_path: Source file/folder path or ID
            target_path: Target folder path or ID
            new_name: Optional new name

        Raises:
            ValueError: If move fails
        """
        logger.info(f"Moving {source_path} to {target_path}")
        try:
            source_id = self._resolve_item_id(source_path)
            target_id = self._resolve_item_id(target_path)

            self.graph.move_item(self.drive_id, source_id, target_id, new_name)
            logger.info(f"Moved: {source_path} to {target_path}")
        except Exception as e:
            logger.error(f"Move failed: {e}")
            raise

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
        """Find item ID from a path.

        Args:
            path: File/folder path (e.g., "Folder/SubFolder/file.pdf")

        Returns:
            Item ID

        Raises:
            ValueError: If path not found
        """
        absolute_path = self._build_remote_path(path)
        item = self.graph.get_item_by_path(self.drive_id, absolute_path)
        return item["id"]

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
