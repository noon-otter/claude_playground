import { useState, useEffect } from 'react';

/**
 * Debug panel that shows console logs directly in the UI
 * Useful when developer tools are hard to access
 */
function DebugPanel() {
  const [logs, setLogs] = useState([]);
  const [isExpanded, setIsExpanded] = useState(true);

  useEffect(() => {
    // Intercept console methods
    const originalLog = console.log;
    const originalError = console.error;
    const originalWarn = console.warn;

    function addLog(message, type = 'log') {
      const timestamp = new Date().toISOString().split('T')[1].slice(0, -1);
      setLogs(prev => [...prev, { message, type, timestamp }].slice(-50)); // Keep last 50 logs
    }

    console.log = function(...args) {
      originalLog.apply(console, args);
      addLog(args.join(' '), 'log');
    };

    console.error = function(...args) {
      originalError.apply(console, args);
      addLog(args.join(' '), 'error');
    };

    console.warn = function(...args) {
      originalWarn.apply(console, args);
      addLog(args.join(' '), 'warn');
    };

    // Restore on unmount
    return () => {
      console.log = originalLog;
      console.error = originalError;
      console.warn = originalWarn;
    };
  }, []);

  if (!isExpanded) {
    return (
      <div style={{
        position: 'fixed',
        bottom: 0,
        left: 0,
        right: 0,
        backgroundColor: '#2d2d2d',
        color: 'white',
        padding: '8px',
        cursor: 'pointer',
        borderTop: '2px solid #0078d4',
        zIndex: 9999
      }} onClick={() => setIsExpanded(true)}>
        <strong>🐛 Debug Console ({logs.length} logs)</strong> - Click to expand
      </div>
    );
  }

  return (
    <div style={{
      position: 'fixed',
      bottom: 0,
      left: 0,
      right: 0,
      backgroundColor: '#1e1e1e',
      color: '#d4d4d4',
      maxHeight: '300px',
      display: 'flex',
      flexDirection: 'column',
      borderTop: '2px solid #0078d4',
      zIndex: 9999,
      fontFamily: 'Consolas, Monaco, monospace',
      fontSize: '11px'
    }}>
      {/* Header */}
      <div style={{
        padding: '8px',
        backgroundColor: '#2d2d2d',
        borderBottom: '1px solid #404040',
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center'
      }}>
        <strong>🐛 Debug Console ({logs.length} logs)</strong>
        <div>
          <button
            onClick={() => setLogs([])}
            style={{
              padding: '4px 8px',
              marginRight: '8px',
              backgroundColor: '#404040',
              color: 'white',
              border: 'none',
              borderRadius: '3px',
              cursor: 'pointer',
              fontSize: '11px'
            }}
          >
            Clear
          </button>
          <button
            onClick={() => setIsExpanded(false)}
            style={{
              padding: '4px 8px',
              backgroundColor: '#404040',
              color: 'white',
              border: 'none',
              borderRadius: '3px',
              cursor: 'pointer',
              fontSize: '11px'
            }}
          >
            Minimize
          </button>
        </div>
      </div>

      {/* Logs */}
      <div style={{
        flex: 1,
        overflowY: 'auto',
        padding: '8px'
      }}>
        {logs.length === 0 ? (
          <div style={{ color: '#666' }}>No logs yet...</div>
        ) : (
          logs.map((log, index) => (
            <div
              key={index}
              style={{
                padding: '4px 0',
                borderBottom: '1px solid #2d2d2d',
                color: log.type === 'error' ? '#f48771' :
                       log.type === 'warn' ? '#dcdcaa' :
                       '#d4d4d4'
              }}
            >
              <span style={{ color: '#666', marginRight: '8px' }}>
                [{log.timestamp}]
              </span>
              <span style={{
                color: log.type === 'error' ? '#f14c4c' :
                       log.type === 'warn' ? '#cca700' :
                       '#4ec9b0',
                marginRight: '8px',
                fontWeight: 'bold'
              }}>
                {log.type.toUpperCase()}
              </span>
              <span>{log.message}</span>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

export default DebugPanel;
