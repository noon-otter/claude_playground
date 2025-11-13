import React from 'react';
import ReactDOM from 'react-dom/client';
import RegisterModal from './taskpane/RegisterModal';
import ErrorBoundary from './components/ErrorBoundary';

console.log('[register.jsx] Script loaded');

// Error fallback UI
function renderError(error) {
  const rootElement = document.getElementById('root');
  if (rootElement) {
    rootElement.innerHTML = `
      <div style="padding: 20px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; margin: 20px;">
        <h2 style="color: #721c24; margin-top: 0;">Failed to Initialize Registration Modal</h2>
        <p><strong>Error:</strong> ${error.message}</p>
        <pre style="background-color: #fff; padding: 10px; border-radius: 4px; overflow: auto; font-size: 12px;">${error.stack}</pre>
        <p style="margin-top: 15px; font-size: 14px;">
          <strong>Debugging tips:</strong><br>
          1. Open Developer Tools: Right-click in the modal and select "Inspect"<br>
          2. Check the Console tab for detailed error messages<br>
          3. Verify your .env file is configured correctly<br>
          4. Make sure the Vite dev server is running on https://localhost:3000
        </p>
        <button onclick="window.location.reload()" style="padding: 8px 16px; background-color: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer;">
          Reload Modal
        </button>
      </div>
    `;
  }
}

// Wait for Office to be ready
console.log('[register.jsx] Waiting for Office.onReady...');

Office.onReady((info) => {
  console.log('[register.jsx] Office.onReady fired!', info);
  console.log('[register.jsx] Host:', info.host);
  console.log('[register.jsx] Platform:', info.platform);

  try {
    console.log('[register.jsx] Creating React root...');
    const rootElement = document.getElementById('root');

    if (!rootElement) {
      throw new Error('Root element not found in DOM');
    }

    const root = ReactDOM.createRoot(rootElement);
    console.log('[register.jsx] Rendering RegisterModal component...');

    root.render(
      <React.StrictMode>
        <ErrorBoundary>
          <RegisterModal />
        </ErrorBoundary>
      </React.StrictMode>
    );

    console.log('[register.jsx] RegisterModal rendered successfully');
  } catch (error) {
    console.error('[register.jsx] Failed to render modal:', error);
    renderError(error);
  }
});
