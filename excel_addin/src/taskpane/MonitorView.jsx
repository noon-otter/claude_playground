import { useState, useEffect } from 'react';
import {
  Card,
  CardHeader,
  Body1,
  Body2,
  Button,
  Badge,
  Divider
} from '@fluentui/react-components';
import {
  Delete24Regular,
  ArrowSync24Regular
} from '@fluentui/react-icons';
import { getModelActivity, removeMonitoredCell } from '../utils/domino-api';

function MonitorView({ modelId, modelConfig }) {
  const [monitoredCells, setMonitoredCells] = useState([]);
  const [recentActivity, setRecentActivity] = useState([]);
  const [isRefreshing, setIsRefreshing] = useState(false);

  useEffect(() => {
    loadMonitoredCells();
    loadActivity();
  }, [modelId]);

  function loadMonitoredCells() {
    if (modelConfig && modelConfig.monitoredCells) {
      setMonitoredCells(modelConfig.monitoredCells);
    }
  }

  async function loadActivity() {
    try {
      const activity = await getModelActivity(modelId, 20);
      setRecentActivity(activity);
    } catch (error) {
      console.error('Failed to load activity:', error);
    }
  }

  async function handleRefresh() {
    setIsRefreshing(true);
    await loadActivity();
    setTimeout(() => setIsRefreshing(false), 500);
  }

  async function handleRemoveCell(cellRange) {
    try {
      await removeMonitoredCell(modelId, cellRange);
      setMonitoredCells(prev => prev.filter(c => c.range !== cellRange));
    } catch (error) {
      console.error('Failed to remove cell:', error);
    }
  }

  function formatTimestamp(timestamp) {
    const date = new Date(timestamp);
    const now = new Date();
    const diffMs = now - date;
    const diffMins = Math.floor(diffMs / 60000);

    if (diffMins < 1) return 'just now';
    if (diffMins < 60) return `${diffMins}m ago`;
    if (diffMins < 1440) return `${Math.floor(diffMins / 60)}h ago`;
    return date.toLocaleDateString();
  }

  return (
    <div>

      {/* Model Info */}
      <Card style={{ marginBottom: '20px' }}>
        <CardHeader
          header={<Body1 weight="semibold">Model Information</Body1>}
        />
        <div style={{ padding: '12px' }}>
          <div style={{ marginBottom: '8px' }}>
            <Body2 style={{ color: '#666' }}>Name:</Body2>
            <Body1>{modelConfig?.name || 'Untitled Model'}</Body1>
          </div>
          <div style={{ marginBottom: '8px' }}>
            <Body2 style={{ color: '#666' }}>Owner:</Body2>
            <Body1>{modelConfig?.owner || 'Unknown'}</Body1>
          </div>
          <div>
            <Body2 style={{ color: '#666' }}>Registered:</Body2>
            <Body1>{modelConfig?.registeredAt ? new Date(modelConfig.registeredAt).toLocaleDateString() : 'Unknown'}</Body1>
          </div>
        </div>
      </Card>

      <Divider style={{ marginBottom: '20px' }} />

      {/* Monitored Cells */}
      <Card style={{ marginBottom: '20px' }}>
        <CardHeader
          header={<Body1 weight="semibold">Monitored Cells</Body1>}
          description={<Body2>{monitoredCells.length} cells tracked</Body2>}
        />
        <div style={{ padding: '12px' }}>
          {monitoredCells.length === 0 ? (
            <Body1 style={{ color: '#999', textAlign: 'center', padding: '20px' }}>
              No cells monitored yet. Use "Mark Input" or "Mark Output" buttons
              in the ribbon to start tracking cells.
            </Body1>
          ) : (
            <div>
              {monitoredCells.map((cell, index) => (
                <div
                  key={index}
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    padding: '12px',
                    marginBottom: '8px',
                    backgroundColor: '#f5f5f5',
                    borderRadius: '4px'
                  }}
                >
                  <div style={{ flex: 1 }}>
                    <Body1 weight="semibold">{cell.range}</Body1>
                    <Body2 style={{ color: '#666' }}>
                      Added {formatTimestamp(cell.addedAt)}
                    </Body2>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Badge
                      appearance="filled"
                      color={cell.type === 'input' ? 'informative' : 'success'}
                    >
                      {cell.type}
                    </Badge>
                    <Button
                      appearance="subtle"
                      icon={<Delete24Regular />}
                      size="small"
                      onClick={() => handleRemoveCell(cell.range)}
                    />
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </Card>

      <Divider style={{ marginBottom: '20px' }} />

      {/* Recent Activity */}
      <Card>
        <CardHeader
          header={<Body1 weight="semibold">Recent Activity</Body1>}
          action={
            <Button
              appearance="subtle"
              icon={<ArrowSync24Regular />}
              size="small"
              onClick={handleRefresh}
              disabled={isRefreshing}
            >
              Refresh
            </Button>
          }
        />
        <div style={{ padding: '12px' }}>
          {recentActivity.length === 0 ? (
            <Body1 style={{ color: '#999', textAlign: 'center', padding: '20px' }}>
              No recent activity
            </Body1>
          ) : (
            <div>
              {recentActivity.slice(0, 10).map((event, index) => (
                <div
                  key={index}
                  style={{
                    padding: '12px',
                    marginBottom: '8px',
                    borderLeft: '3px solid #0078d4',
                    backgroundColor: '#f5f5f5',
                    borderRadius: '2px'
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                    <Body1 weight="semibold">{formatEventType(event.event)}</Body1>
                    <Body2 style={{ color: '#666' }}>{formatTimestamp(event.timestamp)}</Body2>
                  </div>
                  {event.cell && (
                    <Body2 style={{ color: '#666' }}>
                      Cell: {event.cell} {event.value !== undefined && `→ ${event.value}`}
                    </Body2>
                  )}
                  {event.user && (
                    <Body2 style={{ color: '#999', fontSize: '11px' }}>
                      By: {event.user}
                    </Body2>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </Card>

    </div>
  );
}

function formatEventType(eventType) {
  const eventNames = {
    'model_opened': '📂 Model Opened',
    'model_saved': '💾 Model Saved',
    'cell_changed': '✏️ Cell Changed',
    'cell_marked': '🏷️ Cell Marked',
    'selection_changed': '👆 Selection Changed',
    'worksheet_activated': '📄 Worksheet Activated',
    'unmonitored_cell_changed': '📝 Unmonitored Change'
  };

  return eventNames[eventType] || eventType;
}

export default MonitorView;
