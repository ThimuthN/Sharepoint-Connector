"""API routes for Microsoft SharePoint integration."""
import logging
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse
from typing import Optional
from repositories import get_repository
from encryption import TokenEncryption
from oauth import MicrosoftOAuthManager
from token_provider import TokenProvider
from graph_client import MicrosoftGraphClient


logger = logging.getLogger(__name__)

# Lazy load config and create instances
def _get_config():
    from config import get_config
    return get_config()

def _get_manager_and_provider():
    config = _get_config()
    encryption = TokenEncryption(config.TOKEN_ENCRYPTION_KEY)
    oauth_manager = MicrosoftOAuthManager(encryption)
    token_provider = TokenProvider(encryption)
    return oauth_manager, token_provider, config

encryption_instance = None
oauth_manager = None
token_provider = None
config = None

def ensure_initialized():
    global encryption_instance, oauth_manager, token_provider, config
    if oauth_manager is None:
        oauth_manager, token_provider, config = _get_manager_and_provider()
        encryption_instance = TokenEncryption(config.TOKEN_ENCRYPTION_KEY)

repo = get_repository()

router = APIRouter(prefix="/integrations/microsoft", tags=["microsoft"])

# Hardcoded user for MVP (no real auth)
DEMO_USER_ID = "demo_user"


@router.get("/connect")
async def connect():
    """Initiate Microsoft OAuth connection.
    
    Redirects to Microsoft OAuth authorization URL.
    """
    ensure_initialized()
    try:
        auth_url, state = oauth_manager.generate_authorization_url(DEMO_USER_ID)
        return {"auth_url": auth_url, "state": state}
    except Exception as e:
        logger.error(f"OAuth initiation failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/callback")
async def callback(code: str, state: str):
    """Handle Microsoft OAuth callback.
    
    Exchanges authorization code for access token.
    """
    ensure_initialized()
    try:
        connection = oauth_manager.handle_callback(code, state, DEMO_USER_ID)
        return {
            "status": "connected",
            "message": "Successfully connected to Microsoft SharePoint",
            "user_id": connection.microsoft_user_id,
        }
    except ValueError as e:
        logger.error(f"OAuth callback error: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Unexpected callback error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/status")
async def status():
    """Check Microsoft connection status."""
    try:
        connection = repo.get_by_user_id(DEMO_USER_ID)
        if not connection:
            return {"connected": False}
        return {
            "connected": connection.is_connected,
            "microsoft_user_id": connection.microsoft_user_id,
            "expires_at": connection.expires_at.isoformat(),
        }
    except Exception as e:
        logger.error(f"Status check failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/sites")
async def list_sites(search: Optional[str] = Query(None)):
    """List SharePoint sites.
    
    Args:
        search: Optional search query
        
    Returns:
        List of sites
    """
    ensure_initialized()
    try:
        access_token = token_provider.get_valid_access_token(DEMO_USER_ID)
        client = MicrosoftGraphClient(access_token, base_url=config.MICROSOFT_GRAPH_API)
        sites = client.list_sites(search=search)
        return {"value": sites}
    except ValueError as e:
        logger.warning(f"Sites list error: {e}")
        raise HTTPException(status_code=401, detail=str(e))
    except Exception as e:
        logger.error(f"Sites list failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/sites/{site_id}/drives")
async def list_drives(site_id: str):
    """List document libraries for a site.
    
    Args:
        site_id: SharePoint site ID
        
    Returns:
        List of drives
    """
    ensure_initialized()
    try:
        access_token = token_provider.get_valid_access_token(DEMO_USER_ID)
        client = MicrosoftGraphClient(access_token, base_url=config.MICROSOFT_GRAPH_API)
        drives = client.list_drives(site_id)
        return {"value": drives}
    except ValueError as e:
        logger.warning(f"Drives list error: {e}")
        raise HTTPException(status_code=401, detail=str(e))
    except Exception as e:
        logger.error(f"Drives list failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/drives/{drive_id}/items")
async def list_drive_items(
    drive_id: str,
    item_id: Optional[str] = Query(None),
):
    """List items (folders/files) in a drive.
    
    Args:
        drive_id: Drive ID
        item_id: Optional parent item ID (defaults to root)
        
    Returns:
        List of items
    """
    ensure_initialized()
    try:
        access_token = token_provider.get_valid_access_token(DEMO_USER_ID)
        client = MicrosoftGraphClient(access_token, base_url=config.MICROSOFT_GRAPH_API)
        items = client.list_drive_items(drive_id, item_id=item_id)
        return {"value": items}
    except ValueError as e:
        logger.warning(f"Drive items list error: {e}")
        raise HTTPException(status_code=401, detail=str(e))
    except Exception as e:
        logger.error(f"Drive items list failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/drives/{drive_id}/items/{item_id}/download")
async def download_item(drive_id: str, item_id: str):
    """Download a file from a drive.
    
    Args:
        drive_id: Drive ID
        item_id: Item ID of the file
        
    Returns:
        File content as streaming response
    """
    ensure_initialized()
    try:
        access_token = token_provider.get_valid_access_token(DEMO_USER_ID)
        client = MicrosoftGraphClient(access_token, base_url=config.MICROSOFT_GRAPH_API)
        
        # Get item metadata for filename
        item = client._get(f"/drives/{drive_id}/items/{item_id}")
        filename = item.get("name", "download")
        
        # Download file
        content = client.download_drive_item(drive_id, item_id)
        
        return StreamingResponse(
            iter([content]),
            media_type="application/octet-stream",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except ValueError as e:
        logger.warning(f"Download error: {e}")
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Download failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))
