"""Microsoft Graph API client."""
import logging
from typing import Dict, List, Optional
import httpx


logger = logging.getLogger(__name__)


class MicrosoftGraphClient:
    """Client for Microsoft Graph API."""
    
    def __init__(self, access_token: str, base_url: str = "https://graph.microsoft.com/v1.0"):
        self.access_token = access_token
        self.base_url = base_url
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
    
    def get_me(self) -> Dict:
        """Get current user information.
        
        Returns:
            User info dict
            
        Raises:
            ValueError: If request fails
        """
        return self._get("/me")
    
    def list_sites(self, search: Optional[str] = None) -> List[Dict]:
        """List accessible SharePoint sites.
        
        Args:
            search: Optional search query for site name
            
        Returns:
            List of site dicts
            
        Raises:
            ValueError: If request fails
        """
        if search:
            # Search for sites by name
            response = self._get(
                "/sites",
                params={"search": f'"{search}"'},
            )
            return response.get("value", [])
        else:
            # Get all accessible sites
            response = self._get("/sites?search=*")
            return response.get("value", [])
    
    def list_drives(self, site_id: str) -> List[Dict]:
        """List document libraries (drives) for a site.
        
        Args:
            site_id: SharePoint site ID
            
        Returns:
            List of drive dicts
            
        Raises:
            ValueError: If request fails
        """
        response = self._get(f"/sites/{site_id}/drives")
        return response.get("value", [])
    
    def list_drive_items(
        self,
        drive_id: str,
        item_id: Optional[str] = None,
    ) -> List[Dict]:
        """List items (folders/files) in a drive or folder.
        
        Args:
            drive_id: OneDrive/SharePoint drive ID
            item_id: Optional item ID to list children (defaults to root)
            
        Returns:
            List of item dicts
            
        Raises:
            ValueError: If request fails
        """
        if item_id:
            endpoint = f"/drives/{drive_id}/items/{item_id}/children"
        else:
            endpoint = f"/drives/{drive_id}/root/children"
        
        response = self._get(endpoint)
        return response.get("value", [])
    
    def download_drive_item(
        self,
        drive_id: str,
        item_id: str,
    ) -> bytes:
        """Download a file from a drive.
        
        Args:
            drive_id: OneDrive/SharePoint drive ID
            item_id: Item ID of the file
            
        Returns:
            File content as bytes
            
        Raises:
            ValueError: If request fails or item is not a file
        """
        # Get item metadata first to verify it's a file
        item = self._get(f"/drives/{drive_id}/items/{item_id}")
        if "folder" in item:
            raise ValueError("Cannot download folder, only files")
        
        # Download file content
        with httpx.Client() as client:
            response = client.get(
                f"{self.base_url}/drives/{drive_id}/items/{item_id}/content",
                headers=self.headers,
            )
            response.raise_for_status()
            return response.content
    
    def _get(
        self,
        endpoint: str,
        params: Optional[Dict] = None,
    ) -> Dict:
        """Make GET request to Microsoft Graph.
        
        Args:
            endpoint: API endpoint (with leading /)
            params: Optional query parameters
            
        Returns:
            Response JSON as dict
            
        Raises:
            ValueError: If request fails
        """
        url = f"{self.base_url}{endpoint}"
        
        try:
            with httpx.Client() as client:
                response = client.get(
                    url,
                    headers=self.headers,
                    params=params,
                )
                
                if response.status_code == 401:
                    raise ValueError("Unauthorized. Access token may be invalid.")
                elif response.status_code == 403:
                    raise ValueError("Forbidden. Missing required permissions.")
                elif response.status_code == 404:
                    raise ValueError("Not found.")
                elif response.status_code == 429:
                    raise ValueError("Rate limited by Microsoft Graph API.")
                
                response.raise_for_status()
                return response.json()
        except httpx.HTTPError as e:
            logger.error(f"Graph API request failed: {e}")
            raise ValueError(f"Microsoft Graph API error: {str(e)}")
