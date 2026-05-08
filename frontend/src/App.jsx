import React, { useEffect, useState } from 'react';
import MicrosoftConnect from './MicrosoftConnect';
import FileBrowser from './FileBrowser';
import './App.css';

function App() {
  const [refreshKey, setRefreshKey] = useState(0);

  useEffect(() => {
    // Check for OAuth callback
    const params = new URLSearchParams(window.location.search);
    if (params.has('code') || params.has('error')) {
      // OAuth callback occurred, refresh connection status
      setTimeout(() => setRefreshKey((k) => k + 1), 1000);
      // Clean URL
      window.history.replaceState({}, document.title, window.location.pathname);
    }
  }, []);

  return (
    <div className="app">
      <header className="header">
        <h1>Microsoft SharePoint Connector MVP</h1>
      </header>
      
      <main className="main">
        <MicrosoftConnect key={refreshKey} />
        <FileBrowser key={refreshKey} />
      </main>

      <footer className="footer">
        <p>MVP - Delegated user auth only. See README for limitations.</p>
      </footer>
    </div>
  );
}

export default App;
