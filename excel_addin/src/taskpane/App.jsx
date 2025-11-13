import { useState, useEffect } from 'react';
import {
  FluentProvider,
  webLightTheme,
  Title3,
  Body1,
  Badge,
  Button,
  Card,
  CardHeader,
  Divider
} from '@fluentui/react-components';
import {
  CheckmarkCircle24Regular,
  ErrorCircle24Regular,
  Info24Regular
} from '@fluentui/react-icons';
import MonitorView from './MonitorView';
import { loadModel } from '../utils/domino-api';
import { getOrCreateModelId, getWorkbookName } from '../utils/model-id';
import DebugPanel from '../components/DebugPanel';

function App() {
  const [modelId, setModelId] = useState(null);
  const [modelConfig, setModelConfig] = useState(null);
  const [isRegistered, setIsRegistered] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [workbookName, setWorkbookName] = useState('');
  const [error, setError] = useState(null);

  useEffect(() => {
    console.log('[App.jsx] Component mounted, initializing...');
    initializeApp();
  }, []);

  async function initializeApp() {
    console.log('[App.jsx] initializeApp() started');
    try {
      // Get model ID
      console.log('[App.jsx] Getting model ID...');
      const id = await getOrCreateModelId();
      console.log('[App.jsx] Model ID:', id);
      setModelId(id);

      // Get workbook name
      console.log('[App.jsx] Getting workbook name...');
      const name = await getWorkbookName();
      console.log('[App.jsx] Workbook name:', name);
      setWorkbookName(name);

      // Check if registered using architecture-compliant API
      console.log('[App.jsx] Checking if model is registered...');
      try {
        const config = await loadModel(id);
        if (config) {
          console.log('[App.jsx] Model is registered:', config);
          setModelConfig(config);
          setIsRegistered(true);
        } else {
          console.log('[App.jsx] Model is not registered');
        }
      } catch (apiError) {
        // API errors are not critical - the add-in can still work
        console.warn('[App.jsx] API call failed (this is OK in dev mode):', apiError);
      }

      console.log('[App.jsx] Initialization complete');
      setIsLoading(false);
    } catch (error) {
      console.error('[App.jsx] Failed to initialize app:', error);
      setError(error);
      setIsLoading(false);
    }
  }

  function getStatusBadge() {
    if (isLoading) {
      return <Badge appearance="outline" color="subtle">Loading...</Badge>;
    }

    if (isRegistered) {
      return (
        <Badge appearance="filled" color="success" icon={<CheckmarkCircle24Regular />}>
          Active
        </Badge>
      );
    }

    return (
      <Badge appearance="filled" color="warning" icon={<Info24Regular />}>
        Not Registered
      </Badge>
    );
  }

  if (error) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div style={{ padding: '20px' }}>
          <Card style={{ backgroundColor: '#FEF0F0', border: '1px solid #D13438' }}>
            <div style={{ padding: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
                <ErrorCircle24Regular style={{ marginRight: '8px', color: '#D13438' }} />
                <Body1 weight="semibold">Initialization Error</Body1>
              </div>
              <Body1 style={{ marginBottom: '12px', fontSize: '14px' }}>
                {error.message}
              </Body1>
              <details style={{ fontSize: '12px', color: '#666' }}>
                <summary style={{ cursor: 'pointer', marginBottom: '8px' }}>
                  Show stack trace
                </summary>
                <pre style={{
                  backgroundColor: '#f8f9fa',
                  padding: '10px',
                  borderRadius: '4px',
                  overflow: 'auto',
                  fontSize: '11px'
                }}>
                  {error.stack}
                </pre>
              </details>
              <Button
                appearance="primary"
                onClick={() => window.location.reload()}
                style={{ marginTop: '12px' }}
              >
                Reload Add-in
              </Button>
            </div>
          </Card>
        </div>
      </FluentProvider>
    );
  }

  if (isLoading) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div style={{ padding: '20px' }}>
          <Title3>Loading...</Title3>
          <Body1 style={{ marginTop: '10px', color: '#666' }}>
            Initializing Excel add-in...
          </Body1>
        </div>
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={webLightTheme}>
      <DebugPanel />
      <div style={{ padding: '20px', minHeight: '100vh', paddingBottom: '320px' }}>

        {/* Header */}
        <div style={{ marginBottom: '20px' }}>
          <Title3>📊 Domino Governance</Title3>
          <Body1 style={{ color: '#666', marginTop: '8px' }}>
            {workbookName}
          </Body1>
        </div>

        {/* Status Card */}
        <Card style={{ marginBottom: '20px' }}>
          <CardHeader
            header={<Body1 weight="semibold">Monitoring Status</Body1>}
            action={getStatusBadge()}
          />
          <div style={{ padding: '12px' }}>
            <Body1 style={{ fontSize: '12px', color: '#666' }}>
              Model ID: <code style={{ fontSize: '11px' }}>{modelId}</code>
            </Body1>
          </div>
        </Card>

        <Divider style={{ marginBottom: '20px' }} />

        {/* Registration Notice or Monitor View */}
        {!isRegistered ? (
          <Card style={{ backgroundColor: '#FFF4E5', border: '1px solid #FFB900' }}>
            <div style={{ padding: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
                <Info24Regular style={{ marginRight: '8px', color: '#FFB900' }} />
                <Body1 weight="semibold">Model Not Registered</Body1>
              </div>
              <Body1 style={{ marginBottom: '16px', fontSize: '14px' }}>
                This Excel model is not registered with Domino. Click "Register Model"
                in the ribbon to enable governance tracking and monitoring.
              </Body1>
              <Body1 style={{ fontSize: '12px', color: '#666' }}>
                Once registered, all cell changes will be automatically tracked and
                sent to Domino for compliance and audit purposes.
              </Body1>
            </div>
          </Card>
        ) : (
          <MonitorView modelId={modelId} modelConfig={modelConfig} />
        )}

        {/* Footer Info */}
        <div style={{ marginTop: '40px', paddingTop: '20px', borderTop: '1px solid #e0e0e0' }}>
          <Body1 style={{ fontSize: '12px', color: '#999', textAlign: 'center' }}>
            Domino Excel Governance Add-in v1.0
          </Body1>
        </div>

      </div>
    </FluentProvider>
  );
}

export default App;
