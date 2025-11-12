import { useState, useEffect } from 'react';
import {
  FluentProvider,
  webLightTheme,
  Title3,
  Body1,
  Input,
  Textarea,
  Button,
  Field,
  Spinner
} from '@fluentui/react-components';
import { registerModel } from '../utils/domino-api';
import { setModelName, getWorkbookName } from '../utils/model-id';

function RegisterModal() {
  const [modelId, setModelId] = useState('');
  const [modelName, setModelName] = useState('');
  const [owner, setOwner] = useState('');
  const [description, setDescription] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    initializeModal();
  }, []);

  async function initializeModal() {
    try {
      // Get model ID from URL params
      const params = new URLSearchParams(window.location.search);
      const id = params.get('modelId');
      setModelId(id);

      // Pre-fill model name from workbook name
      const workbookName = await getWorkbookName();
      setModelName(workbookName.replace('.xlsx', ''));

      // Try to get user email from Office context
      try {
        const email = Office.context.mailbox?.userProfile?.emailAddress;
        if (email) {
          setOwner(email);
        }
      } catch (e) {
        // Not available in Excel (only Outlook)
        console.log('Email not available from context');
      }
    } catch (error) {
      console.error('Failed to initialize modal:', error);
      setError('Failed to initialize registration form');
    }
  }

  async function handleSubmit(e) {
    e.preventDefault();
    setError(null);

    if (!modelName || !owner) {
      setError('Model name and owner are required');
      return;
    }

    setIsSubmitting(true);

    try {
      // Register with Domino
      const config = await registerModel({
        modelId,
        name: modelName,
        owner,
        description
      });

      // Save model name in document properties
      await setModelName(modelName);

      // Send success message to parent (commands.js)
      Office.context.ui.messageParent(JSON.stringify({
        action: 'registered',
        config
      }));

    } catch (error) {
      console.error('Registration failed:', error);
      setError(error.message || 'Failed to register model');
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
      <div style={{ padding: '20px', maxWidth: '500px', margin: '0 auto' }}>

        <Title3 style={{ marginBottom: '20px' }}>Register Model with Domino</Title3>

        <Body1 style={{ marginBottom: '20px', color: '#666' }}>
          Register this Excel model to enable governance tracking and monitoring.
          All changes to marked cells will be automatically sent to Domino.
        </Body1>

        <form onSubmit={handleSubmit}>

          {/* Model ID (read-only) */}
          <Field
            label="Model ID"
            hint="Unique identifier (auto-generated)"
            style={{ marginBottom: '16px' }}
          >
            <Input
              value={modelId}
              readOnly
              style={{ backgroundColor: '#f5f5f5', fontFamily: 'monospace', fontSize: '12px' }}
            />
          </Field>

          {/* Model Name */}
          <Field
            label="Model Name"
            required
            hint="A descriptive name for this model"
            style={{ marginBottom: '16px' }}
          >
            <Input
              value={modelName}
              onChange={(e) => setModelName(e.target.value)}
              placeholder="e.g., Revenue Forecast 2025"
              required
            />
          </Field>

          {/* Owner */}
          <Field
            label="Owner"
            required
            hint="Primary contact for this model"
            style={{ marginBottom: '16px' }}
          >
            <Input
              type="email"
              value={owner}
              onChange={(e) => setOwner(e.target.value)}
              placeholder="your.email@company.com"
              required
            />
          </Field>

          {/* Description */}
          <Field
            label="Description"
            hint="Optional description of this model's purpose"
            style={{ marginBottom: '24px' }}
          >
            <Textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="Brief description of what this model does..."
              rows={3}
            />
          </Field>

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
              disabled={isSubmitting || !modelName || !owner}
            >
              {isSubmitting ? (
                <>
                  <Spinner size="tiny" style={{ marginRight: '8px' }} />
                  Registering...
                </>
              ) : (
                'Register Model'
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
            <strong>What happens after registration?</strong><br />
            • Model is registered in Domino<br />
            • Background monitoring starts automatically<br />
            • Use ribbon buttons to mark cells as inputs/outputs<br />
            • All changes are streamed to Domino in real-time
          </Body1>
        </div>

      </div>
    </FluentProvider>
  );
}

export default RegisterModal;
