"""Microsoft Graph API client for SharePoint operations."""
import logging
from typing import Dict, List, Optional, BinaryIO
import httpx
from .retry import retry_operation, RetryConfig

logger = logging.getLogger(__name__)


class GraphClient:
    """Client for Microsoft Graph API with retry logic for transient failures."""

    def __init__(self, access_token: str, retry_config: Optional[RetryConfig] = None):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        self.retry_config = retry_config or RetryConfig()

    def get_me(self) -> Dict:
        """Get current user info."""
        return self._get("/me")

    def get_site(self, site_id: str) -> Dict:
        """Get site info."""
        return self._get(f"/sites/{site_id}")

    def list_drives(self, site_id: str) -> List[Dict]:
        """List document libraries (drives) for a site."""
        response = self._get(f"/sites/{site_id}/drives")
        return response.get("value", [])

    def get_drive_item(self, drive_id: str, item_id: str) -> Dict:
        """Get a single drive item (file or folder)."""
        return self._get(f"/drives/{drive_id}/items/{item_id}")

    def list_items(self, drive_id: str, item_id: Optional[str] = None) -> List[Dict]:
        """List items in a folder (or drive root if item_id is None)."""
        if item_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        else:
            endpoint = f"/drives/{drive_id}/root/children"

        response = self._get(endpoint)
        return response.get("value", [])

    def upload_file(
        self,
        drive_id: str,
        file_path: str,
        remote_path: str,
        conflict: str = "overwrite",
    ) -> Dict:
        """Upload a file to SharePoint with conflict handling.

        Args:
            drive_id: Target drive ID
            file_path: Local file path to upload
            remote_path: Remote path (e.g., "Folder/filename.pdf")
            conflict: Conflict behavior:
                - "overwrite": Replace existing file (default for backwards compatibility)
                - "fail_if_exists": Raise error if file exists
                - "rename": Create unique filename if exists (e.g., file (1).pdf)

        Returns:
            Uploaded item metadata

        Raises:
            ValueError: If file exists and conflict="fail_if_exists"
        """
        if conflict not in ("overwrite", "fail_if_exists", "rename"):
            raise ValueError(f"Invalid conflict mode: {conflict}")

        try:
            with open(file_path, "rb") as f:
                file_content = f.read()
        except IOError as e:
            raise ValueError(f"Failed to read file {file_path}: {str(e)}")

        # Split path into folder and filename
        parts = remote_path.split("/")
        filename = parts[-1]
        folder_path = "/".join(parts[:-1]) if len(parts) > 1 else ""

        # Get or create folder structure
        target_item_id = "root"
        if folder_path:
            target_item_id = self._ensure_folder_path(drive_id, folder_path)

        # Handle conflict modes
        if conflict == "fail_if_exists":
            # Check if file exists before upload
            items = self.list_items(drive_id, target_item_id)
            for item in items:
                if item.get("name") == filename and "file" in item:
                    raise ValueError(
                        f"File already exists: {remote_path}. "
                        f"Use conflict='overwrite' to replace or 'rename' to create variant."
                    )

        elif conflict == "rename":
            # Generate unique filename if exists
            items = self.list_items(drive_id, target_item_id)
            existing_names = {item.get("name") for item in items}
            if filename in existing_names:
                # Generate unique name: invoice.pdf → invoice (1).pdf
                filename = self._generate_unique_filename(filename, existing_names)
                logger.info(f"File exists, renaming to: {filename}")

        # Upload file with retry
        url = f"{self.base_url}/drives/{drive_id}/items/{target_item_id}:/{filename}:/content"

        def _do_upload():
            with httpx.Client() as client:
                response = client.put(
                    url,
                    content=file_content,
                    headers={"Authorization": f"Bearer {self.access_token}"},
                )
                response.raise_for_status()
                return response.json()

        try:
            return retry_operation(
                _do_upload,
                self.retry_config,
                operation_name=f"upload {filename}"
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"File upload failed: {e}")
            raise ValueError(f"Failed to upload file: {str(e)}")

    def download_file(self, drive_id: str, item_id: str) -> bytes:
        """Download a file from SharePoint.

        Args:
            drive_id: Source drive ID
            item_id: File item ID

        Returns:
            File content as bytes
        """
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/content"

        def _do_download():
            with httpx.Client() as client:
                response = client.get(
                    url,
                    headers={"Authorization": f"Bearer {self.access_token}"},
                    follow_redirects=True,
                )
                response.raise_for_status()
                return response.content

        try:
            return retry_operation(
                _do_download,
                self.retry_config,
                operation_name=f"download {item_id}"
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"File download failed: {e}")
            raise ValueError(f"Failed to download file: {str(e)}")

    def delete_item(self, drive_id: str, item_id: str) -> None:
        """Delete a file or folder (idempotent).

        Args:
            drive_id: Target drive ID
            item_id: Item ID to delete
        """
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}"

        def _do_delete():
            with httpx.Client() as client:
                response = client.delete(
                    url,
                    headers={"Authorization": f"Bearer {self.access_token}"},
                )
                if response.status_code not in [204, 404]:
                    response.raise_for_status()
                return None

        try:
            return retry_operation(
                _do_delete,
                self.retry_config,
                operation_name=f"delete {item_id}"
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"Item deletion failed: {e}")
            raise ValueError(f"Failed to delete item: {str(e)}")

    def item_exists(self, drive_id: str, item_id: str) -> bool:
        """Check if an item exists."""
        try:
            self.get_drive_item(drive_id, item_id)
            return True
        except:
            return False

    def create_folder(self, drive_id: str, item_id: str, folder_name: str) -> Dict:
        """Create a folder.

        Args:
            drive_id: Target drive ID
            item_id: Parent item ID
            folder_name: Name of folder to create

        Returns:
            Created folder metadata
        """
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/children"
        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename",
        }

        def _do_create():
            with httpx.Client() as client:
                response = client.post(
                    url,
                    json=data,
                    headers=self.headers,
                )
                response.raise_for_status()
                return response.json()

        try:
            return retry_operation(
                _do_create,
                self.retry_config,
                operation_name=f"create folder {folder_name}"
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"Folder creation failed: {e}")
            raise ValueError(f"Failed to create folder: {str(e)}")

    def move_item(
        self,
        drive_id: str,
        item_id: str,
        new_parent_id: str,
        new_name: Optional[str] = None,
    ) -> Dict:
        """Move or rename an item.

        Args:
            drive_id: Drive ID
            item_id: Item to move
            new_parent_id: New parent item ID
            new_name: Optional new name

        Returns:
            Updated item metadata
        """
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}"
        data = {"parentReference": {"id": new_parent_id}}
        if new_name:
            data["name"] = new_name

        def _do_move():
            with httpx.Client() as client:
                response = client.patch(
                    url,
                    json=data,
                    headers=self.headers,
                )
                response.raise_for_status()
                return response.json()

        try:
            return retry_operation(
                _do_move,
                self.retry_config,
                operation_name=f"move {item_id}"
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"Item move failed: {e}")
            raise ValueError(f"Failed to move item: {str(e)}")

    def _ensure_folder_path(self, drive_id: str, folder_path: str) -> str:
        """Ensure folder path exists, creating folders as needed.

        Args:
            drive_id: Target drive ID
            folder_path: Path like "Folder/SubFolder"

        Returns:
            Item ID of the last folder in path
        """
        current_id = "root"
        for folder_name in folder_path.split("/"):
            if not folder_name:
                continue
            try:
                # Try to find existing folder
                items = self.list_items(drive_id, current_id)
                found = None
                for item in items:
                    if item.get("name") == folder_name and "folder" in item:
                        found = item
                        break

                if found:
                    current_id = found["id"]
                else:
                    # Create folder
                    created = self.create_folder(drive_id, current_id, folder_name)
                    current_id = created["id"]
            except Exception as e:
                logger.warning(f"Could not ensure folder {folder_name}: {e}")
                raise

        return current_id

    def _generate_unique_filename(
        self, filename: str, existing_names: set
    ) -> str:
        """Generate a unique filename when target exists.

        Args:
            filename: Original filename (e.g., "invoice.pdf")
            existing_names: Set of existing filenames in target folder

        Returns:
            Unique filename (e.g., "invoice (1).pdf")
        """
        if filename not in existing_names:
            return filename

        # Split into name and extension
        if "." in filename:
            parts = filename.rsplit(".", 1)
            base, ext = parts[0], "." + parts[1]
        else:
            base, ext = filename, ""

        # Find next available number
        counter = 1
        while True:
            candidate = f"{base} ({counter}){ext}"
            if candidate not in existing_names:
                return candidate
            counter += 1

    def _get(self, endpoint: str, params: Optional[Dict] = None) -> Dict:
        """Make GET request to Microsoft Graph with retry.

        Args:
            endpoint: API endpoint (with leading /)
            params: Optional query parameters

        Returns:
            Response JSON as dict
        """
        url = f"{self.base_url}{endpoint}"

        def _do_get():
            with httpx.Client() as client:
                response = client.get(
                    url,
                    headers=self.headers,
                    params=params,
                )

                if response.status_code == 401:
                    raise ValueError("Unauthorized. Access token may be expired.")
                elif response.status_code == 403:
                    raise ValueError("Forbidden. Check folder permissions.")
                elif response.status_code == 404:
                    raise ValueError("Not found.")
                elif response.status_code == 429:
                    raise ValueError("Rate limited. Try again later.")

                response.raise_for_status()
                return response.json()

        try:
            return retry_operation(
                _do_get,
                self.retry_config,
                operation_name=f"GET {endpoint}"
            )
        except httpx.HTTPError as e:
            logger.error(f"Graph API request failed: {e}")
            raise ValueError(f"Graph API error: {str(e)}")
