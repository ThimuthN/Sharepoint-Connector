import React, { useEffect, useState } from 'react';
import api from './api';
import './App.css';

function FileBrowser() {
  const [connected, setConnected] = useState(false);
  const [sites, setSites] = useState([]);
  const [selectedSite, setSelectedSite] = useState(null);
  const [drives, setDrives] = useState([]);
  const [selectedDrive, setSelectedDrive] = useState(null);
  const [items, setItems] = useState([]);
  const [breadcrumb, setBreadcrumb] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [searchQuery, setSearchQuery] = useState('');

  useEffect(() => {
    checkStatusAndLoadSites();
  }, []);

  const checkStatusAndLoadSites = async () => {
    try {
      const status = await api.getStatus();
      setConnected(status.connected);
      if (status.connected) {
        await loadSites();
      }
    } catch (err) {
      setError('Failed to check connection status');
    }
  };

  const loadSites = async (search = null) => {
    setLoading(true);
    setError(null);
    try {
      const data = await api.listSites(search);
      setSites(data.value || []);
    } catch (err) {
      setError('Failed to load sites');
    } finally {
      setLoading(false);
    }
  };

  const handleSiteSelect = async (site) => {
    setSelectedSite(site);
    setDrives([]);
    setItems([]);
    setBreadcrumb([]);
    setSelectedDrive(null);
    setLoading(true);
    setError(null);
    
    try {
      const data = await api.listDrives(site.id);
      setDrives(data.value || []);
    } catch (err) {
      setError('Failed to load drives');
    } finally {
      setLoading(false);
    }
  };

  const handleDriveSelect = async (drive) => {
    setSelectedDrive(drive);
    setBreadcrumb([]);
    setLoading(true);
    setError(null);
    
    try {
      const data = await api.listItems(drive.id);
      setItems(data.value || []);
    } catch (err) {
      setError('Failed to load items');
    } finally {
      setLoading(false);
    }
  };

  const handleFolderOpen = async (item) => {
    setLoading(true);
    setError(null);
    
    try {
      const data = await api.listItems(selectedDrive.id, item.id);
      setItems(data.value || []);
      setBreadcrumb([...breadcrumb, item]);
    } catch (err) {
      setError('Failed to load folder contents');
    } finally {
      setLoading(false);
    }
  };

  const handleBreadcrumbClick = async (index) => {
    if (index === -1) {
      // Go back to drive root
      const data = await api.listItems(selectedDrive.id);
      setItems(data.value || []);
      setBreadcrumb([]);
      return;
    }
    
    const item = breadcrumb[index];
    const newBreadcrumb = breadcrumb.slice(0, index);
    
    try {
      const data = await api.listItems(selectedDrive.id, item.id);
      setItems(data.value || []);
      setBreadcrumb(newBreadcrumb);
    } catch (err) {
      setError('Failed to load folder');
    }
  };

  const handleDownload = async (item) => {
    setError(null);
    try {
      await api.downloadItem(selectedDrive.id, item.id, item.name);
    } catch (err) {
      setError('Failed to download item');
    }
  };

  const handleSearch = () => {
    if (searchQuery.trim()) {
      loadSites(searchQuery);
    }
  };

  if (!connected) {
    return (
      <div className="browser-section">
        <p className="info">Connect your Microsoft account to browse SharePoint</p>
      </div>
    );
  }

  return (
    <div className="browser-section">
      <h2>SharePoint File Browser</h2>
      {error && <div className="error">{error}</div>}
      
      {loading && <div className="loading">Loading...</div>}

      {!selectedSite ? (
        <div>
          <h3>SharePoint Sites</h3>
          <div className="search-bar">
            <input
              type="text"
              placeholder="Search sites..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              onKeyPress={(e) => e.key === 'Enter' && handleSearch()}
            />
            <button onClick={handleSearch} className="btn btn-small">
              Search
            </button>
          </div>
          <ul className="item-list">
            {sites.map((site) => (
              <li key={site.id} className="site-item">
                <button
                  onClick={() => handleSiteSelect(site)}
                  className="item-button"
                >
                  {site.displayName}
                </button>
              </li>
            ))}
          </ul>
          {sites.length === 0 && !loading && (
            <p className="info">No sites found</p>
          )}
        </div>
      ) : !selectedDrive ? (
        <div>
          <button onClick={() => setSelectedSite(null)} className="btn btn-small">
            ← Back to Sites
          </button>
          <h3>{selectedSite.displayName}</h3>
          <h4>Document Libraries</h4>
          <ul className="item-list">
            {drives.map((drive) => (
              <li key={drive.id} className="drive-item">
                <button
                  onClick={() => handleDriveSelect(drive)}
                  className="item-button"
                >
                  📚 {drive.name}
                </button>
              </li>
            ))}
          </ul>
          {drives.length === 0 && !loading && (
            <p className="info">No document libraries found</p>
          )}
        </div>
      ) : (
        <div>
          <button
            onClick={() => {
              setSelectedDrive(null);
              setBreadcrumb([]);
              setItems([]);
            }}
            className="btn btn-small"
          >
            ← Back to Libraries
          </button>
          <h3>{selectedDrive.name}</h3>
          
          <div className="breadcrumb">
            <button onClick={() => handleBreadcrumbClick(-1)} className="breadcrumb-item">
              Root
            </button>
            {breadcrumb.map((item, index) => (
              <React.Fragment key={item.id}>
                <span>/</span>
                <button
                  onClick={() => handleBreadcrumbClick(index)}
                  className="breadcrumb-item"
                >
                  {item.name}
                </button>
              </React.Fragment>
            ))}
          </div>

          <ul className="item-list">
            {items.map((item) => (
              <li key={item.id} className="file-item">
                <span className="item-icon">
                  {item.folder ? '📁' : '📄'}
                </span>
                <span className="item-name">{item.name}</span>
                {item.folder ? (
                  <button
                    onClick={() => handleFolderOpen(item)}
                    className="btn btn-tiny"
                  >
                    Open
                  </button>
                ) : (
                  <button
                    onClick={() => handleDownload(item)}
                    className="btn btn-tiny"
                  >
                    Download
                  </button>
                )}
              </li>
            ))}
          </ul>
          {items.length === 0 && !loading && (
            <p className="info">Folder is empty</p>
          )}
        </div>
      )}
    </div>
  );
}

export default FileBrowser;
