import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

// Office Initialize with fallback
const initializeApp = () => {
  const root = ReactDOM.createRoot(
    document.getElementById('root') as HTMLElement
  );

  root.render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
};

if (typeof Office !== 'undefined') {
  console.log('Office.js loaded, waiting for ready...');
  Office.onReady((reason) => {
    console.log('Office Add-in ready:', reason);
    initializeApp();
  });
} else {
  console.log('Office.js not available, starting app anyway...');
  // Fallback: start app immediately if Office.js isn't available
  initializeApp();
}
