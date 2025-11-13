/*
 * commands-v2.js - Architecture-Compliant Background Script
 * Implements the exact event flow from DEPLOYMENT.md
 *
 * Event Flows:
 * 1. On File Load → GET /wb/load-model (restore tracked ranges)
 * 2. Register Model → PUT /wb/upsert-model (create with version=1)
 * 3. Update Model → PUT /wb/upsert-model (increment version)
 * 4. Tracked Range Change → POST /wb/create-model-trace
 */

// Global state
let monitoringActive = false;
let modelConfig = null;  // WorkbookModel: {model_name, tracked_ranges[], model_id, version}
let traceQueue = [];
let isOnline = true;
let currentUsername = 'unknown';

// Domino API endpoint - UPDATE THIS
const DOMINO_API_BASE = 'http://localhost:5000';

/**
 * Initialize on Office ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('🟢 Domino Governance Add-in loaded (Architecture-Compliant v2)');
    initializeMonitoring();
  }
});

// =====================================================
// EVENT FLOW 1: On File Load (Workbook Load Event)
// =====================================================

/**
 * Main initialization - checks if model is registered and starts monitoring
 */
async function initializeMonitoring() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;

      // Get or create persistent model ID
      const modelId = await getOrCreateModelId(workbook, context);
      console.log(`📋 Model ID: ${modelId}`);

      // Get current username
      currentUsername = getUserEmail();
      console.log(`👤 Username: ${currentUsername}`);

      // EVENT: Workbook Load
      // Call: GET /wb/load-model
      const registered = await loadModelFromBackend(modelId);

      if (registered) {
        modelConfig = registered;
        await startLiveMonitoring(workbook, context, modelId);
        console.log(`✅ Monitoring active for "${registered.model_name}" v${registered.version}`);
      } else {
        console.log(`ℹ️ Model not registered. Use "Register Model" button to enable monitoring.`);
      }

      await context.sync();
    });
  } catch (error) {
    console.error('Failed to initialize monitoring:', error);
  }
}

/**
 * GET /wb/load-model
 * Load model metadata by model_id
 */
async function loadModelFromBackend(modelId) {
  try {
    const response = await fetch(
      `${DOMINO_API_BASE}/wb/load-model?model_id=${encodeURIComponent(modelId)}`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.ok) {
      const model = await response.json();
      console.log('[loadModelFromBackend] Model loaded:', model);
      return model;
    }

    if (response.status === 404) {
      console.log('[loadModelFromBackend] Model not found (not registered yet)');
      return null;
    }

    console.error('[loadModelFromBackend] Failed:', response.statusText);
    return null;
  } catch (error) {
    console.error('[loadModelFromBackend] Error:', error);
    return null;
  }
}

// =====================================================
// EVENT FLOW 2: User-Driven: Register Model
// =====================================================

/**
 * Ribbon command: Show registration modal
 */
async function showRegisterModal() {
  try {
    const modelId = await getModelIdForModal();

    Office.context.ui.displayDialogAsync(
      `https://localhost:3000/register.html?modelId=${modelId}`,
      { height: 60, width: 40, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const dialog = result.value;

          // Handle messages from modal
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            const message = JSON.parse(arg.message);

            if (message.action === 'registered') {
              // Reload monitoring with new config
              modelConfig = message.config;
              console.log('✅ Model registered successfully:', modelConfig);

              // Restart monitoring
              initializeMonitoring();
            }

            dialog.close();
          });

          // Handle dialog closed
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.log('Dialog closed:', arg.error);
          });
        }
      }
    );
  } catch (error) {
    console.error('Failed to show registration modal:', error);
  }
}

// =====================================================
// EVENT FLOW 3: User-Driven: Update Model
// =====================================================
// Tracked ranges are managed through the Registration Modal
// Users can update tracked ranges by re-opening the modal

// =====================================================
// EVENT FLOW 4: Event-Driven: On Tracked Range Changes
// =====================================================

/**
 * Start live monitoring of all events
 */
async function startLiveMonitoring(workbook, context, modelId) {
  if (monitoringActive) {
    console.log('⚠️ Monitoring already active');
    return;
  }

  // Monitor ALL worksheet changes
  workbook.worksheets.onChanged.add(async (event) => {
    await handleCellChange(event, modelId);
  });

  monitoringActive = true;
  console.log('🔴 Live monitoring started');
}

/**
 * Handle cell change events
 */
async function handleCellChange(event, modelId) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      const range = sheet.getRange(event.address);

      range.load(['values', 'address']);
      await context.sync();

      // Check if this cell is in a tracked range
      const trackedRange = findTrackedRange(event.address);

      if (trackedRange) {
        // EVENT: Tracked Range Change
        // Call: POST /wb/create-model-trace
        await createTrace({
          model_id: modelId,
          timestamp: new Date().toISOString(),
          tracked_range_name: trackedRange.name,
          username: currentUsername,
          value: range.values[0][0]
        });

        console.log(
          `📝 Trace: ${trackedRange.name} = ${range.values[0][0]} by ${currentUsername}`
        );
      }
    });
  } catch (error) {
    console.error('Error handling cell change:', error);
  }
}

/**
 * Find tracked range that contains the given address
 */
function findTrackedRange(address) {
  if (!modelConfig || !modelConfig.tracked_ranges) {
    return null;
  }

  // Simple containment check - can be enhanced for complex ranges
  return modelConfig.tracked_ranges.find(tr => {
    return address.includes(tr.range) || tr.range.includes(address);
  });
}

/**
 * POST /wb/create-model-trace
 * Create a trace log entry
 */
async function createTrace(traceData) {
  try {
    const response = await fetch(`${DOMINO_API_BASE}/wb/create-model-trace`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(traceData),
      keepalive: true // Survives page navigation
    });

    if (response.ok) {
      isOnline = true;
      // Flush queue if we have offline traces
      flushTraceQueue();
      return true;
    } else {
      console.warn(`[createTrace] API returned ${response.status}`);
      return false;
    }
  } catch (error) {
    // Queue trace if offline
    console.warn('[createTrace] API unreachable, queuing trace:', error.message);
    isOnline = false;
    queueTraceLocally(traceData);
    return false;
  }
}

// =====================================================
// OFFLINE QUEUE MANAGEMENT
// =====================================================

/**
 * Queue trace locally when offline
 */
function queueTraceLocally(traceData) {
  traceQueue.push({
    ...traceData,
    queuedAt: new Date().toISOString()
  });

  // Limit queue size
  if (traceQueue.length > 100) {
    traceQueue.shift(); // Remove oldest
  }

  // Try to flush periodically
  setTimeout(flushTraceQueue, 30000); // Retry in 30s
}

/**
 * Flush queued traces when back online
 */
async function flushTraceQueue() {
  if (traceQueue.length === 0 || !isOnline) {
    return;
  }

  console.log(`📤 Flushing ${traceQueue.length} queued traces`);

  const traces = [...traceQueue];
  traceQueue = [];

  try {
    const response = await fetch(`${DOMINO_API_BASE}/wb/create-model-trace-batch`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model_id: modelConfig.model_id,
        timestamp: new Date().toISOString(),
        changes: traces.map(t => ({
          tracked_range_name: t.tracked_range_name,
          value: t.value
        })),
        username: currentUsername
      })
    });

    if (!response.ok) {
      // Re-queue on failure
      traceQueue = traces.concat(traceQueue);
    }
  } catch (error) {
    // Re-queue on failure
    traceQueue = traces.concat(traceQueue);
    isOnline = false;
  }
}

// =====================================================
// UTILITY FUNCTIONS
// =====================================================

/**
 * Get or create persistent model ID from custom document properties
 */
async function getOrCreateModelId(workbook, context) {
  const properties = workbook.properties.custom;
  properties.load('items');
  await context.sync();

  // Check if model ID exists
  let modelIdProp = null;
  for (let i = 0; i < properties.items.length; i++) {
    if (properties.items[i].key === 'DominoModelId') {
      modelIdProp = properties.items[i];
      break;
    }
  }

  if (modelIdProp) {
    modelIdProp.load('value');
    await context.sync();
    return modelIdProp.value;
  }

  // Generate new persistent ID
  const modelId = generateModelId();
  properties.add('DominoModelId', modelId);
  await context.sync();

  return modelId;
}

/**
 * Generate unique model ID
 */
function generateModelId() {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 9);
  return `excel_${timestamp}_${random}`;
}

/**
 * Get model ID for modal
 */
async function getModelIdForModal() {
  let modelId = null;

  await Excel.run(async (context) => {
    const workbook = context.workbook;
    modelId = await getOrCreateModelId(workbook, context);
    await context.sync();
  });

  return modelId;
}

/**
 * Get current user identity from Office 365
 * Tries multiple methods to extract user email/identity
 */
function getUserEmail() {
  try {
    // Method 1: Office.context.mailbox (Outlook only)
    if (Office.context.mailbox?.userProfile?.emailAddress) {
      return Office.context.mailbox.userProfile.emailAddress;
    }

    // Method 2: Try to get from Office platform info
    // This is a fallback - we'll try to get it asynchronously later
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.auth) {
      // Trigger async token fetch
      getUserEmailAsync().then(email => {
        if (email && email !== 'unknown') {
          currentUsername = email;
          console.log(`✅ Updated username from token: ${email}`);
        }
      }).catch(err => {
        console.warn('Could not fetch user identity:', err);
      });
    }

    // Method 3: Try Office.context.document info
    const contextUser = Office.context?.document?.url;
    if (contextUser && contextUser.includes('@')) {
      // Sometimes the URL contains user info
      const match = contextUser.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
      if (match) {
        return match[1];
      }
    }

    // Fallback: use 'anonymous' as placeholder
    return 'anonymous';
  } catch (error) {
    console.warn('Error getting user email:', error);
    return 'anonymous';
  }
}

/**
 * Async method to get user email from Office 365 access token
 */
async function getUserEmailAsync() {
  try {
    if (typeof OfficeRuntime === 'undefined' || !OfficeRuntime.auth) {
      return null;
    }

    // Get access token (requires Office 365)
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: true
    });

    // Parse JWT to extract email (simple parsing - not validating signature)
    const base64Url = token.split('.')[1];
    const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
    const jsonPayload = decodeURIComponent(atob(base64).split('').map(function(c) {
      return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
    }).join(''));

    const payload = JSON.parse(jsonPayload);

    // Extract email from token claims
    const email = payload.upn || payload.email || payload.unique_name || payload.preferred_username;

    if (email) {
      console.log(`🔐 Got user identity from token: ${email}`);
      return email;
    }

    return null;
  } catch (error) {
    console.warn('Could not get access token:', error.message);
    return null;
  }
}

// =====================================================
// REGISTER RIBBON COMMAND HANDLERS
// =====================================================

Office.actions.associate("showRegisterModal", showRegisterModal);

console.log('📋 Command handlers registered');
