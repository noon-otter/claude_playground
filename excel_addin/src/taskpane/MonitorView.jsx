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
import { getModelTraces, upsertModel } from '../utils/domino-api';

function MonitorView({ modelId, modelConfig }) {
  const [trackedRanges, setTrackedRanges] = useState([]);
  const [recentTraces, setRecentTraces] = useState([]);
  const [isRefreshing, setIsRefreshing] = useState(false);

  useEffect(() => {
    loadTrackedRanges();
    loadTraces();
  }, [modelId]);

  function loadTrackedRanges() {
    // Use tracked_ranges from the new architecture
    if (modelConfig && modelConfig.tracked_ranges) {
      setTrackedRanges(modelConfig.tracked_ranges);
    }
  }

  async function loadTraces() {
    try {
      const traces = await getModelTraces(modelId, 20);
      setRecentTraces(traces);
    } catch (error) {
      console.error('Failed to load traces:', error);
    }
  }

  async function handleRefresh() {
    setIsRefreshing(true);
    await loadTraces();
    setTimeout(() => setIsRefreshing(false), 500);
  }

  async function handleRemoveRange(rangeName) {
    try {
      // Remove tracked range by updating the model
      const updatedRanges = trackedRanges.filter(tr => tr.name !== rangeName);

      await upsertModel({
        model_name: modelConfig.model_name,
        tracked_ranges: updatedRanges,
        model_id: modelId,
        version: modelConfig.version
      });

      setTrackedRanges(updatedRanges);
    } catch (error) {
      console.error('Failed to remove tracked range:', error);
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
            <Body1>{modelConfig?.model_name || 'Untitled Model'}</Body1>
          </div>
          <div style={{ marginBottom: '8px' }}>
            <Body2 style={{ color: '#666' }}>Version:</Body2>
            <Body1>v{modelConfig?.version || 1}</Body1>
          </div>
          <div>
            <Body2 style={{ color: '#666' }}>Model ID:</Body2>
            <Body1 style={{ fontSize: '11px', fontFamily: 'monospace' }}>
              {modelConfig?.model_id || modelId}
            </Body1>
          </div>
        </div>
      </Card>

      <Divider style={{ marginBottom: '20px' }} />

      {/* Tracked Ranges */}
      <Card style={{ marginBottom: '20px' }}>
        <CardHeader
          header={<Body1 weight="semibold">Tracked Ranges</Body1>}
          description={<Body2>{trackedRanges.length} ranges tracked</Body2>}
        />
        <div style={{ padding: '12px' }}>
          {trackedRanges.length === 0 ? (
            <Body1 style={{ color: '#999', textAlign: 'center', padding: '20px' }}>
              No ranges tracked yet. Use "Add Tracked Range" button
              in the ribbon to start tracking cell ranges.
            </Body1>
          ) : (
            <div>
              {trackedRanges.map((trackedRange, index) => (
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
                    <Body1 weight="semibold">{trackedRange.name}</Body1>
                    <Body2 style={{ color: '#666', fontFamily: 'monospace', fontSize: '11px' }}>
                      {trackedRange.range}
                    </Body2>
                  </div>
                  <Button
                    appearance="subtle"
                    icon={<Delete24Regular />}
                    size="small"
                    onClick={() => handleRemoveRange(trackedRange.name)}
                  />
                </div>
              ))}
            </div>
          )}
        </div>
      </Card>

      <Divider style={{ marginBottom: '20px' }} />

      {/* Recent Traces */}
      <Card>
        <CardHeader
          header={<Body1 weight="semibold">Recent Traces</Body1>}
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
          {recentTraces.length === 0 ? (
            <Body1 style={{ color: '#999', textAlign: 'center', padding: '20px' }}>
              No traces recorded yet
            </Body1>
          ) : (
            <div>
              {recentTraces.slice(0, 10).map((trace, index) => (
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
                    <Body1 weight="semibold">{trace.tracked_range_name}</Body1>
                    <Body2 style={{ color: '#666' }}>{formatTimestamp(trace.timestamp)}</Body2>
                  </div>
                  <Body2 style={{ color: '#666' }}>
                    Value: {JSON.stringify(trace.value)}
                  </Body2>
                  {trace.username && (
                    <Body2 style={{ color: '#999', fontSize: '11px' }}>
                      By: {trace.username}
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

export default MonitorView;
