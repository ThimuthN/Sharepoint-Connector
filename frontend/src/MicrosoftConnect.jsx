import React, { useEffect, useState } from 'react';
import api from './api';
import './App.css';

function MicrosoftConnect() {
  const [connected, setConnected] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [userId, setUserId] = useState(null);

  useEffect(() => {
    checkStatus();
  }, []);

  const checkStatus = async () => {
    try {
      const status = await api.getStatus();
      setConnected(status.connected);
      if (status.connected) {
        setUserId(status.microsoft_user_id);
      }
    } catch (err) {
      console.error('Status check failed:', err);
    }
  };

  const handleConnect = async () => {
    setLoading(true);
    setError(null);
    try {
      const { auth_url } = await api.getConnectUrl();
      window.location.href = auth_url;
    } catch (err) {
      setError('Failed to initiate connection. Check configuration.');
      setLoading(false);
    }
  };

  const handleDisconnect = async () => {
    setConnected(false);
    setUserId(null);
    // In a real app, would call disconnect endpoint
  };

  return (
    <div className="connect-section">
      <h2>Microsoft SharePoint Connection</h2>
      {error && <div className="error">{error}</div>}
      
      {connected ? (
        <div className="connected">
          <p className="status-badge">✓ Connected</p>
          <p>User: {userId}</p>
          <button onClick={handleDisconnect} className="btn btn-secondary">
            Disconnect
          </button>
        </div>
      ) : (
        <button
          onClick={handleConnect}
          disabled={loading}
          className="btn btn-primary"
        >
          {loading ? 'Connecting...' : 'Connect Microsoft Account'}
        </button>
      )}
    </div>
  );
}

export default MicrosoftConnect;
