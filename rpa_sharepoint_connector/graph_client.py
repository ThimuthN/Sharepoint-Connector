"""Microsoft Graph API client for SharePoint operations."""
import logging
import os
import tempfile
from typing import Dict, List, Optional
from urllib.parse import quote

import httpx

from .retry import retry_operation, RetryConfig

logger = logging.getLogger(__name__)


class GraphClient:
    """Client for Microsoft Graph API with retry logic for transient failures."""

    # Keep simple uploads small because this path buffers the full payload in memory.
    SIMPLE_UPLOAD_LIMIT_BYTES = 4 * 1024 * 1024
    UPLOAD_SESSION_CHUNK_BYTES = 10 * 1024 * 1024  # 10 MiB, multiple of 320 KiB
    REQUEST_TIMEOUT_SECONDS = 30.0

    def __init__(self, access_token: str, retry_config: Optional[RetryConfig] = None):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        self.retry_config = retry_config or RetryConfig()
        self._client: Optional[httpx.Client] = None

    def close(self) -> None:
        """Close the underlying HTTP client if it was created."""
        if self._client is not None:
            self._client.close()
            self._client = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()

    def _get_client(self) -> httpx.Client:
        """Lazily create and reuse one HTTP client per GraphClient instance."""
        if self._client is None:
            self._client = httpx.Client()
        return self._client

    def get_me(self) -> Dict:
        """Get current user info."""
        return self._get("/me")

    def get_site(self, site_id: str) -> Dict:
        """Get site info."""
        return self._get(f"/sites/{site_id}")

    def list_drives(self, site_id: str) -> List[Dict]:
        """List document libraries (drives) for a site."""
        return self._get_paginated(f"/sites/{site_id}/drives")

    def get_drive_item(self, drive_id: str, item_id: str) -> Dict:
        """Get a single drive item (file or folder)."""
        return self._get(f"/drives/{drive_id}/items/{item_id}")

    def get_item_by_path(self, drive_id: str, item_path: str) -> Dict:
        """Get a drive item by path relative to drive root."""
        normalized_path = self._normalize_drive_path(item_path)
        if not normalized_path:
            return self._get(f"/drives/{drive_id}/root")

        encoded_path = "/".join(
            quote(segment, safe="")
            for segment in normalized_path.split("/")
        )
        return self._get(f"/drives/{drive_id}/root:/{encoded_path}")

    def list_items(self, drive_id: str, item_id: Optional[str] = None) -> List[Dict]:
        """List items in a folder (or drive root if item_id is None)."""
        if item_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        else:
            endpoint = f"/drives/{drive_id}/root/children"

        return self._get_paginated(endpoint)

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

        file_size = None
        try:
            file_size = os.path.getsize(file_path)
        except OSError:
            # Keep legacy testability/paths where open() may be mocked without a real file.
            file_size = None

        # Split path into folder and filename
        parts = remote_path.split("/")
        filename = parts[-1]
        folder_path = "/".join(parts[:-1]) if len(parts) > 1 else ""

        # Get or create folder structure
        target_item_id = "root"
        if folder_path:
            target_item_id = self._ensure_folder_path(drive_id, folder_path)

        # Preserve current deterministic conflict behavior for small and large uploads.
        if conflict == "fail_if_exists":
            items = self.list_items(drive_id, target_item_id)
            for item in items:
                if item.get("name") == filename and "file" in item:
                    raise ValueError(
                        f"File already exists: {remote_path}. "
                        "Use conflict='overwrite' to replace or 'rename' to create variant."
                    )
        elif conflict == "rename":
            items = self.list_items(drive_id, target_item_id)
            existing_names = {item.get("name") for item in items}
            if filename in existing_names:
                filename = self._generate_unique_filename(filename, existing_names)
                logger.info(f"File exists, renaming to: {filename}")

        if file_size is not None and file_size > self.SIMPLE_UPLOAD_LIMIT_BYTES:
            return self._upload_file_via_session(
                drive_id=drive_id,
                target_item_id=target_item_id,
                file_path=file_path,
                filename=filename,
                conflict=conflict,
                file_size=file_size,
            )

        return self._upload_file_simple(
            drive_id=drive_id,
            target_item_id=target_item_id,
            file_path=file_path,
            filename=filename,
        )

    def _upload_file_simple(
        self,
        drive_id: str,
        target_item_id: str,
        file_path: str,
        filename: str,
    ) -> Dict:
        """Upload a file with simple PUT /content API."""
        try:
            with open(file_path, "rb") as f:
                file_content = f.read()
        except OSError as e:
            raise ValueError(f"Failed to read file {file_path}: {str(e)}") from e

        url = f"{self.base_url}/drives/{drive_id}/items/{target_item_id}:/{filename}:/content"

        def _do_upload():
            client = self._get_client()
            response = client.put(
                url,
                content=file_content,
                headers={"Authorization": f"Bearer {self.access_token}"},
                timeout=self.REQUEST_TIMEOUT_SECONDS,
            )
            response.raise_for_status()
            return response.json()

        try:
            return retry_operation(
                _do_upload,
                self.retry_config,
                operation_name=f"upload {filename}",
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"File upload failed: {e}")
            raise ValueError(f"Failed to upload file: {str(e)}") from e

    def _upload_file_via_session(
        self,
        drive_id: str,
        target_item_id: str,
        file_path: str,
        filename: str,
        conflict: str,
        file_size: int,
    ) -> Dict:
        """Upload large files via Graph upload session."""
        if file_size <= 0:
            raise ValueError("Cannot upload empty file.")

        conflict_map = {
            "overwrite": "replace",
            "fail_if_exists": "fail",
            "rename": "rename",
        }
        session_url = (
            f"{self.base_url}/drives/{drive_id}/items/{target_item_id}:/{filename}:/createUploadSession"
        )
        payload = {
            "item": {
                "name": filename,
                "@microsoft.graph.conflictBehavior": conflict_map[conflict],
            }
        }

        client = self._get_client()
        session_response = client.post(
            session_url,
            json=payload,
            headers=self.headers,
            timeout=self.REQUEST_TIMEOUT_SECONDS,
        )
        session_response.raise_for_status()
        upload_url = session_response.json().get("uploadUrl")
        if not upload_url:
            raise ValueError("Upload session created but no uploadUrl returned.")

        chunk_size = self.UPLOAD_SESSION_CHUNK_BYTES
        if chunk_size % (320 * 1024) != 0:
            raise ValueError("Upload chunk size must be a multiple of 320 KiB.")

        with open(file_path, "rb") as f:
            start = 0
            while start < file_size:
                end = min(start + chunk_size, file_size) - 1
                chunk = f.read(end - start + 1)
                if not chunk:
                    raise ValueError("Unexpected EOF while reading upload chunk.")

                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                }

                def _do_chunk_upload():
                    response = client.put(
                        upload_url,
                        content=chunk,
                        headers=headers,
                        timeout=self.REQUEST_TIMEOUT_SECONDS,
                    )
                    response.raise_for_status()
                    return response

                chunk_response = retry_operation(
                    _do_chunk_upload,
                    self.retry_config,
                    operation_name=f"upload chunk {start}-{end}",
                )

                if chunk_response.status_code in (200, 201):
                    return chunk_response.json()

                if chunk_response.status_code != 202:
                    raise ValueError(
                        f"Unexpected chunk upload status {chunk_response.status_code}"
                    )

                start = end + 1

        raise ValueError("Large file upload did not complete successfully.")

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
            client = self._get_client()
            response = client.get(
                url,
                headers={"Authorization": f"Bearer {self.access_token}"},
                follow_redirects=True,
                timeout=self.REQUEST_TIMEOUT_SECONDS,
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
            raise ValueError(f"Failed to download file: {str(e)}") from e

    def download_file_to_path(self, drive_id: str, item_id: str, local_path: str) -> None:
        """Download a file directly to disk with retry-safe temp-file staging."""
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/content"
        target_path = os.path.abspath(local_path)
        target_dir = os.path.dirname(target_path) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)

        def _do_download_to_path():
            temp_path = None
            try:
                with tempfile.NamedTemporaryFile(
                    mode="wb",
                    delete=False,
                    dir=target_dir,
                    prefix=".download_",
                    suffix=".tmp",
                ) as tmp_file:
                    temp_path = tmp_file.name

                client = self._get_client()
                with client.stream(
                    "GET",
                    url,
                    headers={"Authorization": f"Bearer {self.access_token}"},
                    follow_redirects=True,
                    timeout=self.REQUEST_TIMEOUT_SECONDS,
                ) as response:
                    response.raise_for_status()
                    with open(temp_path, "wb") as output_file:
                        for chunk in response.iter_bytes():
                            if chunk:
                                output_file.write(chunk)

                os.replace(temp_path, target_path)
                return None
            except Exception:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.unlink(temp_path)
                    except OSError:
                        pass
                raise

        try:
            return retry_operation(
                _do_download_to_path,
                self.retry_config,
                operation_name=f"download {item_id} to file",
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"File download failed: {e}")
            raise ValueError(f"Failed to download file: {str(e)}") from e

    def delete_item(self, drive_id: str, item_id: str) -> None:
        """Delete a file or folder (idempotent).

        Args:
            drive_id: Target drive ID
            item_id: Item ID to delete
        """
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}"

        def _do_delete():
            client = self._get_client()
            response = client.delete(
                url,
                headers={"Authorization": f"Bearer {self.access_token}"},
                timeout=self.REQUEST_TIMEOUT_SECONDS,
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
            raise ValueError(f"Failed to delete item: {str(e)}") from e

    def item_exists(self, drive_id: str, item_id: str) -> bool:
        """Check if an item exists."""
        try:
            self.get_drive_item(drive_id, item_id)
            return True
        except ValueError as exc:
            if "not found" in str(exc).lower():
                return False
            raise
        except Exception:
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
            client = self._get_client()
            response = client.post(
                url,
                json=data,
                headers=self.headers,
                timeout=self.REQUEST_TIMEOUT_SECONDS,
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
            raise ValueError(f"Failed to create folder: {str(e)}") from e

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
            client = self._get_client()
            response = client.patch(
                url,
                json=data,
                headers=self.headers,
                timeout=self.REQUEST_TIMEOUT_SECONDS,
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
            raise ValueError(f"Failed to move item: {str(e)}") from e

    def _ensure_folder_path(self, drive_id: str, folder_path: str) -> str:
        """Ensure folder path exists, creating folders as needed.

        Args:
            drive_id: Target drive ID
            folder_path: Path like "Folder/SubFolder"

        Returns:
            Item ID of the last folder in path
        """
        normalized_path = self._normalize_drive_path(folder_path)
        if not normalized_path:
            return "root"

        try:
            existing = self.get_item_by_path(drive_id, normalized_path)
            if "folder" not in existing:
                raise ValueError(f"Path exists but is not a folder: {normalized_path}")
            return existing["id"]
        except ValueError as exc:
            if "not found" not in str(exc).lower():
                raise

        current_id = "root"
        current_path_parts = []
        for folder_name in normalized_path.split("/"):
            current_path_parts.append(folder_name)
            current_path = "/".join(current_path_parts)
            try:
                existing = self.get_item_by_path(drive_id, current_path)
                if "folder" not in existing:
                    raise ValueError(f"Path exists but is not a folder: {current_path}")
                current_id = existing["id"]
            except Exception as e:
                if "not found" not in str(e).lower():
                    logger.warning(f"Could not ensure folder {folder_name}: {e}")
                    raise
                created = self.create_folder(drive_id, current_id, folder_name)
                current_id = created["id"]

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

    @staticmethod
    def _normalize_drive_path(path: str) -> str:
        """Normalize drive-relative path by removing empty segments."""
        return "/".join(segment for segment in (path or "").split("/") if segment)

    def _get_json(self, url: str, operation_name: str, params: Optional[Dict] = None) -> Dict:
        """GET a JSON payload with shared Graph status handling and retry."""
        def _do_get():
            client = self._get_client()
            response = client.get(
                url,
                headers=self.headers,
                params=params,
                timeout=self.REQUEST_TIMEOUT_SECONDS,
            )

            if response.status_code == 401:
                raise ValueError("Unauthorized. Access token may be expired.")
            if response.status_code == 403:
                raise ValueError("Forbidden. Check folder permissions.")
            if response.status_code == 404:
                raise ValueError("Not found.")
            if response.status_code == 429:
                request = getattr(response, "request", httpx.Request("GET", url))
                raise httpx.HTTPStatusError(
                    "Rate limited. Try again later.",
                    request=request,
                    response=response,
                )

            response.raise_for_status()
            return response.json()

        try:
            return retry_operation(
                _do_get,
                self.retry_config,
                operation_name=operation_name,
            )
        except httpx.HTTPError as e:
            logger.error(f"Graph API request failed: {e}")
            raise ValueError(f"Graph API error: {str(e)}") from e

    def _get(self, endpoint: str, params: Optional[Dict] = None) -> Dict:
        """Make GET request to Microsoft Graph with retry.

        Args:
            endpoint: API endpoint (with leading /)
            params: Optional query parameters

        Returns:
            Response JSON as dict
        """
        url = f"{self.base_url}{endpoint}"
        return self._get_json(url, operation_name=f"GET {endpoint}", params=params)

    def _get_paginated(self, endpoint: str) -> List[Dict]:
        """Fetch all pages for list endpoints using @odata.nextLink."""
        first_page = self._get(endpoint)
        items = self._get_page_items(first_page, endpoint)
        next_link = first_page.get("@odata.nextLink")
        seen_links = set()

        while next_link:
            if not isinstance(next_link, str):
                raise ValueError("Graph API returned invalid @odata.nextLink value.")
            if next_link in seen_links:
                raise ValueError(f"Detected pagination loop while fetching {endpoint}.")
            seen_links.add(next_link)

            page = self._get_absolute(next_link)
            items.extend(self._get_page_items(page, endpoint))
            next_link = page.get("@odata.nextLink")

        return items

    @staticmethod
    def _get_page_items(page: Dict, endpoint: str) -> List[Dict]:
        """Validate paginated payload shape before aggregating items."""
        if not isinstance(page, dict):
            raise ValueError(f"Graph API returned invalid page payload while fetching {endpoint}.")

        items = page.get("value", [])
        if not isinstance(items, list):
            raise ValueError(f"Graph API returned invalid item list while fetching {endpoint}.")

        return list(items)

    def _get_absolute(self, url: str) -> Dict:
        """GET an absolute URL with retry (used for nextLink paging)."""
        return self._get_json(url, operation_name="GET next page")
