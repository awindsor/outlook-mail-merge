import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

// Office Initialize
if (typeof Office !== 'undefined') {
  Office.onReady(() => {
    console.log('Office Add-in ready');
  });
}

const root = ReactDOM.createRoot(
  document.getElementById('root') as HTMLElement
);

root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
