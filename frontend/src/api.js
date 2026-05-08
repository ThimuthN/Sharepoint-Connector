"""Frontend API client for Microsoft integration."""
const API_BASE_URL = process.env.REACT_APP_API_URL || 'http://localhost:8000/api';

class MicrosoftIntegrationApi {
  async getConnectUrl() {
    const response = await fetch(`${API_BASE_URL}/integrations/microsoft/connect`);
    if (!response.ok) throw new Error('Failed to get connect URL');
    return response.json();
  }

  async getStatus() {
    const response = await fetch(`${API_BASE_URL}/integrations/microsoft/status`);
    if (!response.ok) throw new Error('Failed to get status');
    return response.json();
  }

  async listSites(search = null) {
    const url = new URL(`${API_BASE_URL}/integrations/microsoft/sites`);
    if (search) url.searchParams.append('search', search);
    
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to list sites');
    return response.json();
  }

  async listDrives(siteId) {
    const response = await fetch(
      `${API_BASE_URL}/integrations/microsoft/sites/${siteId}/drives`
    );
    if (!response.ok) throw new Error('Failed to list drives');
    return response.json();
  }

  async listItems(driveId, itemId = null) {
    const url = new URL(`${API_BASE_URL}/integrations/microsoft/drives/${driveId}/items`);
    if (itemId) url.searchParams.append('item_id', itemId);
    
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to list items');
    return response.json();
  }

  async downloadItem(driveId, itemId, filename) {
    const response = await fetch(
      `${API_BASE_URL}/integrations/microsoft/drives/${driveId}/items/${itemId}/download`
    );
    if (!response.ok) throw new Error('Failed to download item');
    
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
  }
}

export default new MicrosoftIntegrationApi();
