import React from 'react';
import { createRoot } from 'react-dom/client';
import App from '../components/App';

/* global document, Office */

const initializeApp = () => {
  const container = document.getElementById('container');
  if (container) {
    const root = createRoot(container);
    root.render(<App />);
  }
};

// Check if Office.js is available (Excel environment)
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      initializeApp();
    }
  });
} else {
  // Fallback for development environment (browser)
  document.addEventListener('DOMContentLoaded', initializeApp);
}