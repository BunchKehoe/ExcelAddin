import React from 'react';
import { createRoot } from 'react-dom/client';
import App from '../components/App';

/* global document, Office, CustomFunctions */

// Type declaration for CustomFunctions
declare const CustomFunctions: {
  associate: (id: string, func: Function) => void;
} | undefined;

// Import custom functions for shared runtime support
import { IRR, JOINCELLS } from '../functions/functions';

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
      // For shared runtime - register custom functions in taskpane context
      if (typeof CustomFunctions !== 'undefined') {
        CustomFunctions.associate("PC.IRR", IRR);
        CustomFunctions.associate("PC.JOINCELLS", JOINCELLS);
        console.log('Custom functions registered in shared runtime (taskpane context)');
      }
      
      initializeApp();
    }
  });
} else {
  // Fallback for development environment (browser)
  document.addEventListener('DOMContentLoaded', initializeApp);
}