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
import { getModelById } from '../utils/domino-api';
import { getOrCreateModelId, getWorkbookName } from '../utils/model-id';

function App() {
  const [modelId, setModelId] = useState(null);
  const [modelConfig, setModelConfig] = useState(null);
  const [isRegistered, setIsRegistered] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [workbookName, setWorkbookName] = useState('');

  useEffect(() => {
    initializeApp();
  }, []);

  async function initializeApp() {
    try {
      // Get model ID
      const id = await getOrCreateModelId();
      setModelId(id);

      // Get workbook name
      const name = await getWorkbookName();
      setWorkbookName(name);

      // Check if registered
      const config = await getModelById(id);
      if (config) {
        setModelConfig(config);
        setIsRegistered(true);
      }

      setIsLoading(false);
    } catch (error) {
      console.error('Failed to initialize app:', error);
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

  if (isLoading) {
    return (
      <FluentProvider theme={webLightTheme}>
        <div style={{ padding: '20px' }}>
          <Title3>Loading...</Title3>
        </div>
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={webLightTheme}>
      <div style={{ padding: '20px', minHeight: '100vh' }}>

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
