import { useState, useEffect } from 'react';
import {
  FluentProvider,
  webLightTheme,
  Title3,
  Body1,
  Input,
  Button,
  Field,
  Spinner,
  Card
} from '@fluentui/react-components';
import { Delete24Regular, Add24Regular } from '@fluentui/react-icons';
import { upsertModel } from '../utils/domino-api';
import { setModelName, getWorkbookName } from '../utils/model-id';
import DebugPanel from '../components/DebugPanel';

function RegisterModal() {
  const [modelId, setModelId] = useState('');
  const [modelName, setModelName] = useState('');
  const [trackedRanges, setTrackedRanges] = useState([]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    console.log('[RegisterModal] Initializing...');
    initializeModal();
  }, []);

  async function initializeModal() {
    try {
      // Get model ID from URL params
      const params = new URLSearchParams(window.location.search);
      const id = params.get('modelId');
      console.log('[RegisterModal] Model ID:', id);
      setModelId(id);

      // Try to load existing model
      try {
        const response = await fetch(`http://localhost:5000/wb/load-model?model_id=${id}`);
        if (response.ok) {
          const model = await response.json();
          console.log('[RegisterModal] Loaded existing model:', model);
          setModelName(model.model_name);
          setTrackedRanges(model.tracked_ranges || []);
        } else {
          // Model not registered yet, pre-fill model name from workbook name
          const workbookName = await getWorkbookName();
          console.log('[RegisterModal] Workbook name:', workbookName);
          setModelName(workbookName.replace(/\.xlsx?$/i, ''));
        }
      } catch (loadError) {
        // Model not registered yet, pre-fill model name from workbook name
        const workbookName = await getWorkbookName();
        console.log('[RegisterModal] Workbook name:', workbookName);
        setModelName(workbookName.replace(/\.xlsx?$/i, ''));
      }

      console.log('[RegisterModal] Ready');
    } catch (error) {
      console.error('[RegisterModal] Initialization failed:', error);
      setError(`Failed to initialize: ${error.message}`);
    }
  }

  function addTrackedRange() {
    setTrackedRanges([...trackedRanges, { name: '', range: '' }]);
  }

  function updateTrackedRange(index, field, value) {
    const updated = [...trackedRanges];
    updated[index][field] = value;
    setTrackedRanges(updated);
  }

  function deleteTrackedRange(index) {
    const updated = trackedRanges.filter((_, i) => i !== index);
    setTrackedRanges(updated);
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setError(null);

    if (!modelName.trim()) {
      setError('Model name is required');
      return;
    }

    // Validate tracked ranges
    const validRanges = trackedRanges.filter(r => r.name.trim() && r.range.trim());
    if (validRanges.length !== trackedRanges.length) {
      setError('Please fill in all tracked range fields or remove empty ones');
      return;
    }

    setIsSubmitting(true);

    try {
      console.log('[RegisterModal] Registering model...');

      // Register model with tracked ranges
      const config = await upsertModel({
        model_name: modelName,
        tracked_ranges: validRanges,
        model_id: modelId
      });
      console.log('[RegisterModal] Model registered:', config);

      // Save model name in document properties
      await setModelName(modelName);

      // Send success message to parent (commands.js)
      console.log('[RegisterModal] Notifying parent...');
      Office.context.ui.messageParent(JSON.stringify({
        action: 'registered',
        config
      }));

    } catch (error) {
      console.error('[RegisterModal] Registration failed:', error);
      setError(error.message || 'Failed to register model. Check console for details.');
      setIsSubmitting(false);
    }
  }

  function handleCancel() {
    Office.context.ui.messageParent(JSON.stringify({
      action: 'cancelled'
    }));
  }

  return (
    <FluentProvider theme={webLightTheme}>
      <DebugPanel />
      <div style={{ padding: '20px', maxWidth: '600px', margin: '0 auto', paddingBottom: '100px' }}>

        <Title3 style={{ marginBottom: '20px' }}>Register Model</Title3>

        <Body1 style={{ marginBottom: '20px', color: '#666' }}>
          Register this workbook as a tracked model. Define which cell ranges to monitor for changes.
        </Body1>

        <form onSubmit={handleSubmit}>

          {/* Model ID (read-only) */}
          <Field
            label="Model ID"
            hint="Unique identifier (auto-generated, read-only)"
            style={{ marginBottom: '16px' }}
          >
            <Input
              value={modelId}
              readOnly
              style={{
                backgroundColor: '#f5f5f5',
                fontFamily: 'monospace',
                fontSize: '12px',
                color: '#666'
              }}
            />
          </Field>

          {/* Model Name */}
          <Field
            label="Model Name"
            required
            hint="A descriptive name for this model"
            style={{ marginBottom: '24px' }}
          >
            <Input
              value={modelName}
              onChange={(e) => setModelName(e.target.value)}
              placeholder="e.g., Revenue Forecast 2025"
              required
            />
          </Field>

          {/* Tracked Ranges */}
          <div style={{ marginBottom: '24px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
              <Body1 style={{ fontWeight: 600 }}>Tracked Ranges</Body1>
              <Button
                appearance="subtle"
                icon={<Add24Regular />}
                onClick={addTrackedRange}
                disabled={isSubmitting}
              >
                Add Range
              </Button>
            </div>

            <Body1 style={{ fontSize: '12px', color: '#666', marginBottom: '12px' }}>
              Define cell ranges to monitor. Changes to these ranges will be tracked.
            </Body1>

            {trackedRanges.length === 0 && (
              <Card style={{ padding: '20px', textAlign: 'center', backgroundColor: '#f9f9f9' }}>
                <Body1 style={{ color: '#999' }}>
                  No tracked ranges yet. Click "Add Range" to start monitoring cells.
                </Body1>
              </Card>
            )}

            {trackedRanges.map((range, index) => (
              <Card key={index} style={{ padding: '12px', marginBottom: '12px' }}>
                <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
                  <div style={{ flex: 1 }}>
                    <Field
                      label="Name"
                      size="small"
                      style={{ marginBottom: '8px' }}
                    >
                      <Input
                        value={range.name}
                        onChange={(e) => updateTrackedRange(index, 'name', e.target.value)}
                        placeholder="e.g., Revenue, Assumptions"
                        size="small"
                        disabled={isSubmitting}
                      />
                    </Field>
                  </div>
                  <div style={{ flex: 1 }}>
                    <Field
                      label="Range"
                      size="small"
                      style={{ marginBottom: '8px' }}
                    >
                      <Input
                        value={range.range}
                        onChange={(e) => updateTrackedRange(index, 'range', e.target.value)}
                        placeholder="e.g., Sheet1!A1:D10"
                        size="small"
                        disabled={isSubmitting}
                      />
                    </Field>
                  </div>
                  <Button
                    appearance="subtle"
                    icon={<Delete24Regular />}
                    onClick={() => deleteTrackedRange(index)}
                    disabled={isSubmitting}
                    style={{ marginTop: '20px' }}
                    title="Remove this range"
                  />
                </div>
              </Card>
            ))}
          </div>

          {/* Error message */}
          {error && (
            <div style={{
              padding: '12px',
              marginBottom: '16px',
              backgroundColor: '#FDE7E9',
              border: '1px solid #C50F1F',
              borderRadius: '4px'
            }}>
              <Body1 style={{ color: '#C50F1F' }}>{error}</Body1>
            </div>
          )}

          {/* Actions */}
          <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
            <Button
              appearance="secondary"
              onClick={handleCancel}
              disabled={isSubmitting}
            >
              Cancel
            </Button>
            <Button
              type="submit"
              appearance="primary"
              disabled={isSubmitting || !modelName.trim()}
            >
              {isSubmitting ? (
                <>
                  <Spinner size="tiny" style={{ marginRight: '8px' }} />
                  Saving...
                </>
              ) : (
                'Save & Register'
              )}
            </Button>
          </div>

        </form>

        {/* Info box */}
        <div style={{
          marginTop: '24px',
          padding: '12px',
          backgroundColor: '#F3F2F1',
          borderRadius: '4px'
        }}>
          <Body1 style={{ fontSize: '12px', color: '#666' }}>
            <strong>After registration:</strong><br />
            • Model is saved with a unique ID<br />
            • Changes to tracked ranges are automatically logged<br />
            • You can update tracked ranges later by re-opening this dialog
          </Body1>
        </div>

      </div>
    </FluentProvider>
  );
}

export default RegisterModal;
