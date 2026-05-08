"""Simple SDK for RPA bots to interact with SharePoint."""
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Optional
from .auth import MicrosoftAuth
from .token_store import TokenStore
from .graph_client import GraphClient
from .profiles import ProfileManager

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

    def __init__(self, profile: str, store_dir: Optional[str] = None):
        """Initialize SharePoint client.

        Args:
            profile: Profile name (must be configured first)
            store_dir: Optional token store directory

        Raises:
            ValueError: If profile not found or tokens invalid
        """
        self.store = TokenStore(store_dir=store_dir)
        self.profile_manager = ProfileManager(self.store)

        # Load profile
        profile_data = self.store.load_profile(profile)
        if not profile_data:
            raise ValueError(
                f"Profile '{profile}' not found. "
                f"Run: python -m rpa_sharepoint_connector configure --profile {profile}"
            )

        self.profile = profile
        self.profile_data = profile_data
        self.drive_id = profile_data["drive_id"]
        self.folder_id = profile_data["folder_id"]

        # Initialize client (uses Innobot public-client app)
        self.auth = MicrosoftAuth()
        self._ensure_valid_token()
        self.graph = GraphClient(self.profile_data["access_token"])

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
            # If remote_path looks like a path, find the item ID first
            if "/" in remote_path:
                item_id = self._find_item_id(remote_path)
            else:
                item_id = remote_path

            content = self.graph.download_file(self.drive_id, item_id)
            with open(local_path, "wb") as f:
                f.write(content)
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
            if "/" in remote_path:
                item_id = self._find_item_id(remote_path)
            else:
                item_id = remote_path

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
            if "/" in remote_path:
                item_id = self._find_item_id(remote_path)
            else:
                item_id = remote_path
            return self.graph.item_exists(self.drive_id, item_id)
        except:
            return False

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
            parts = folder_path.split("/")
            parent_id = self.folder_id

            for folder_name in parts:
                if not folder_name:
                    continue
                result = self.graph.create_folder(self.drive_id, parent_id, folder_name)
                parent_id = result["id"]

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
            source_id = (
                self._find_item_id(source_path)
                if "/" in source_path
                else source_path
            )
            target_id = (
                self._find_item_id(target_path) if "/" in target_path else target_path
            )

            self.graph.move_item(self.drive_id, source_id, target_id, new_name)
            logger.info(f"Moved: {source_path} to {target_path}")
        except Exception as e:
            logger.error(f"Move failed: {e}")
            raise

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
        current_id = self.folder_id
        for part in path.split("/"):
            if not part:
                continue
            items = self.graph.list_items(self.drive_id, current_id)
            found = None
            for item in items:
                if item.get("name") == part:
                    found = item
                    break

            if not found:
                raise ValueError(f"Path not found: {path}")
            current_id = found["id"]

        return current_id

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
