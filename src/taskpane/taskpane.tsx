import React from 'react';
import { createRoot } from 'react-dom/client';
import App from '../components/App';

/* global document, Office, CustomFunctions */

// Initialize custom functions
const initializeCustomFunctions = async () => {
  // Import and register custom functions
  if (typeof CustomFunctions !== 'undefined') {
    // Import the functions module
    await import('../functions/functions');
    console.log('Custom functions loaded and registered');
  }
};

const initializeApp = () => {
  const container = document.getElementById('container');
  if (container) {
    const root = createRoot(container);
    root.render(<App />);
  }
};

// Check if Office.js is available (Excel environment)
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
      // Initialize custom functions first
      await initializeCustomFunctions();
      // Then initialize the app
      initializeApp();
    }
  });
} else {
  // Fallback for development environment (browser)
  document.addEventListener('DOMContentLoaded', initializeApp);
}