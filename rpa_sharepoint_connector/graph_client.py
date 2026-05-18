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
    # Increased to 100MB to use reliable simple uploads instead of session-based
    SIMPLE_UPLOAD_LIMIT_BYTES = 100 * 1024 * 1024  # 100 MB simple upload limit
    UPLOAD_SESSION_CHUNK_BYTES = 5 * 1024 * 1024  # 5 MiB for session chunks (unused now)
    REQUEST_TIMEOUT_SECONDS = 300.0  # 5 minutes for large file uploads

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
        """Upload a file to SharePoint with conflict handling."""
        try:
            file_size = os.path.getsize(file_path)
        except OSError as e:
            raise ValueError(f"Failed to get file size: {str(e)}") from e

        normalized_path = self._normalize_drive_path(remote_path)
        if "/" in normalized_path:
            folder_path = "/".join(normalized_path.split("/")[:-1])
            filename = normalized_path.split("/")[-1]
            target_item_id = self._ensure_folder_path(drive_id, folder_path) if folder_path else "root"
        else:
            filename = normalized_path
            target_item_id = "root"

        if file_size <= self.SIMPLE_UPLOAD_LIMIT_BYTES:
            return self._upload_file_simple(drive_id, target_item_id, file_path, filename)
        else:
            return self._upload_file_via_session(drive_id, target_item_id, file_path, filename, conflict, file_size)

    def _upload_file_simple(
        self,
        drive_id: str,
        target_item_id: str,
        file_path: str,
        filename: str,
    ) -> Dict:
        """Upload a file with simple PUT /content API."""
        # Validate file exists and is not a symlink
        if not os.path.isfile(file_path) or os.path.islink(file_path):
            raise ValueError(f"Invalid file path (not a regular file or is symlink): {file_path}")

        # Validate file size before reading
        try:
            file_size = os.path.getsize(file_path)
            if file_size > self.SIMPLE_UPLOAD_LIMIT_BYTES:
                raise ValueError(f"File exceeds simple upload limit ({file_size} > {self.SIMPLE_UPLOAD_LIMIT_BYTES} bytes)")
        except OSError as e:
            raise ValueError(f"Failed to get file size: {str(e)}") from e

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
        # Use simple approach: upload directly to root, let Graph handle conflicts
        # Encode filename for URL
        encoded_filename = quote(filename, safe="")
        session_url = (
            f"{self.base_url}/drives/{drive_id}/root:/{encoded_filename}:/createUploadSession"
        )
        payload = {
            "item": {
                "name": filename,
                "@microsoft.graph.conflictBehavior": conflict_map[conflict],
            }
        }

        def _create_upload_session():
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
            return upload_url

        try:
            upload_url = retry_operation(
                _create_upload_session,
                self.retry_config,
                operation_name=f"create upload session for {filename}",
            )
        except (ValueError, httpx.HTTPError) as e:
            logger.error(f"Failed to create upload session: {e}")
            raise ValueError(f"Failed to create upload session: {str(e)}") from e

        chunk_size = self.UPLOAD_SESSION_CHUNK_BYTES
        if chunk_size % (320 * 1024) != 0:
            raise ValueError("Upload chunk size must be a multiple of 320 KiB.")

        with open(file_path, "rb") as f:
            start = 0
            chunk_num = 1
            while start < file_size:
                end = min(start + chunk_size, file_size) - 1
                chunk = f.read(end - start + 1)
                if not chunk:
                    raise ValueError("Unexpected EOF while reading upload chunk.")

                percent_complete = int((end + 1) / file_size * 100)
                size_mb = file_size / (1024 * 1024)
                logger.info(f"Uploading chunk {chunk_num} ({percent_complete}% complete, {size_mb:.2f} MB total)")

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
                    operation_name=f"upload chunk {chunk_num} ({percent_complete}%)",
                )

                if chunk_response.status_code in (200, 201):
                    logger.info(f"Large file upload completed: {filename}")
                    try:
                        return chunk_response.json()
                    except (ValueError, RuntimeError):
                        # Response has no JSON body, but upload succeeded
                        return {"id": "", "name": filename}

                if chunk_response.status_code == 204:
                    logger.info(f"Large file upload completed: {filename}")
                    return {"id": "", "name": filename}

                if chunk_response.status_code != 202:
                    raise ValueError(
                        f"Unexpected chunk upload status {chunk_response.status_code}"
                    )

                start = end + 1
                chunk_num += 1

        raise ValueError("Large file upload did not complete successfully.")

    def download_file(self, drive_id: str, item_id: str) -> bytes:
        """Download a file from SharePoint."""
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

        # Resolve real path to prevent symlink attacks
        target_path = os.path.realpath(target_path)
        target_dir = os.path.dirname(target_path) or os.getcwd()

        # Validate target directory is not a symlink
        if os.path.islink(target_dir):
            raise ValueError(f"Target directory is a symlink (security risk): {target_dir}")

        # Check parent directory too (prevent traversal to parent via symlink)
        parent_dir = os.path.dirname(target_dir)
        if parent_dir and os.path.islink(parent_dir):
            raise ValueError(f"Parent directory is a symlink (security risk): {parent_dir}")

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
        """Delete a file or folder (idempotent)."""
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
        """Create a folder."""
        url = f"{self.base_url}/drives/{drive_id}/items/{item_id}/children"
        data = {"name": folder_name, "folder": {}}

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
        """Move or rename an item."""
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
        """Ensure folder path exists, creating folders as needed."""
        if not folder_path or folder_path == "/":
            return "root"

        normalized = self._normalize_drive_path(folder_path)
        current_item_id = "root"

        for folder_name in normalized.split("/"):
            try:
                item = self.get_item_by_path(drive_id, f"{'/'.join(normalized.split('/')[:normalized.split('/').index(folder_name)+1])}")
                current_item_id = item["id"]
            except ValueError:
                created = self.create_folder(drive_id, current_item_id, folder_name)
                current_item_id = created["id"]

        return current_item_id

    def _generate_unique_filename(
        self, filename: str, existing_names: set
    ) -> str:
        """Generate a unique filename when target exists."""
        if filename not in existing_names:
            return filename

        name_parts = filename.rsplit(".", 1)
        base = name_parts[0]
        ext = f".{name_parts[1]}" if len(name_parts) > 1 else ""

        counter = 1
        while True:
            new_name = f"{base} ({counter}){ext}"
            if new_name not in existing_names:
                return new_name
            counter += 1
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
        """Make GET request to Microsoft Graph with retry."""
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
