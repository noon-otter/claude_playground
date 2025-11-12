/*
 * commands.js - ALWAYS RUNNING BACKGROUND SCRIPT
 * This runs independently of the taskpane being open
 * Handles all event monitoring and streaming to Domino
 */

// Global state
let monitoringActive = false;
let modelConfig = null;
let eventQueue = [];
let isOnline = true;

// Domino API endpoint - UPDATE THIS
const DOMINO_API_BASE = 'https://your-domino.com/api';

/**
 * Initialize on Office ready
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log('🟢 Domino Governance Add-in loaded');
    initializeMonitoring();
  }
});

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

      // Check if this model is registered in Domino
      const registered = await checkModelRegistration(modelId);

      if (registered) {
        modelConfig = registered;
        await startLiveMonitoring(workbook, context, modelId);
        console.log(`✅ Monitoring active for "${registered.name}"`);

        // Send "opened" event
        streamToDomino('model_opened', {
          modelId,
          timestamp: new Date().toISOString()
        });
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
 * Check if model is registered in Domino
 */
async function checkModelRegistration(modelId) {
  try {
    const response = await fetch(`${DOMINO_API_BASE}/models/${modelId}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      }
    });

    if (response.ok) {
      return await response.json();
    }

    return null;
  } catch (error) {
    console.error('Failed to check registration:', error);
    return null;
  }
}

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

  // Monitor selection changes (user interactions)
  workbook.onSelectionChanged.add((event) => {
    streamToDomino('selection_changed', {
      modelId,
      worksheet: event.worksheetId,
      range: event.address,
      timestamp: new Date().toISOString()
    });
  });

  // Monitor saves
  workbook.onAutoSaveSettingChanged.add(() => {
    streamToDomino('model_saved', {
      modelId,
      timestamp: new Date().toISOString()
    });
  });

  // Monitor worksheet activations
  workbook.worksheets.onActivated.add((event) => {
    streamToDomino('worksheet_activated', {
      modelId,
      worksheetId: event.worksheetId,
      timestamp: new Date().toISOString()
    });
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

      range.load(['values', 'formulas', 'address']);
      await context.sync();

      // Check if this cell is in our monitored list
      const isMonitored = checkIfMonitored(event.address);

      if (isMonitored) {
        const cellType = getCellType(event.address);

        streamToDomino('cell_changed', {
          modelId,
          worksheet: event.worksheetId,
          cell: event.address,
          value: range.values[0][0],
          formula: range.formulas[0][0],
          type: cellType, // 'input' or 'output'
          user: getUserEmail(),
          timestamp: new Date().toISOString()
        });
      } else {
        // Still log unmonitored changes for full audit trail
        streamToDomino('unmonitored_cell_changed', {
          modelId,
          worksheet: event.worksheetId,
          cell: event.address,
          timestamp: new Date().toISOString()
        });
      }
    });
  } catch (error) {
    console.error('Error handling cell change:', error);
  }
}

/**
 * Check if a cell range is in monitored list
 */
function checkIfMonitored(address) {
  if (!modelConfig || !modelConfig.monitoredCells) {
    return false;
  }

  return modelConfig.monitoredCells.some(cell => {
    return cellInRange(address, cell.range);
  });
}

/**
 * Get cell type (input/output) from monitored list
 */
function getCellType(address) {
  if (!modelConfig || !modelConfig.monitoredCells) {
    return null;
  }

  const cell = modelConfig.monitoredCells.find(c => cellInRange(address, c.range));
  return cell ? cell.type : null;
}

/**
 * Check if address is within a range
 */
function cellInRange(address, range) {
  // Simple check - can be enhanced for complex ranges
  return address.includes(range) || range.includes(address);
}

/**
 * Get current user email
 */
function getUserEmail() {
  try {
    // Try to get from Office context
    if (Office.context.mailbox && Office.context.mailbox.userProfile) {
      return Office.context.mailbox.userProfile.emailAddress;
    }
    return 'unknown';
  } catch {
    return 'unknown';
  }
}

/**
 * Stream event to Domino API
 */
function streamToDomino(eventType, data) {
  const payload = {
    event: eventType,
    ...data
  };

  // Non-blocking fire-and-forget
  fetch(`${DOMINO_API_BASE}/excel-events`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload),
    keepalive: true // Survives page navigation
  })
    .then(response => {
      if (response.ok) {
        isOnline = true;
        // Flush queue if we have offline events
        flushEventQueue();
      } else {
        console.warn(`API returned ${response.status}`);
      }
    })
    .catch(error => {
      // Queue event if offline
      console.warn('Domino API unreachable, queuing event:', error.message);
      isOnline = false;
      queueEventLocally(eventType, data);
    });
}

/**
 * Queue event locally when offline
 */
function queueEventLocally(eventType, data) {
  eventQueue.push({
    event: eventType,
    ...data,
    queuedAt: new Date().toISOString()
  });

  // Limit queue size
  if (eventQueue.length > 100) {
    eventQueue.shift(); // Remove oldest
  }

  // Try to flush periodically
  setTimeout(flushEventQueue, 30000); // Retry in 30s
}

/**
 * Flush queued events when back online
 */
async function flushEventQueue() {
  if (eventQueue.length === 0 || !isOnline) {
    return;
  }

  console.log(`📤 Flushing ${eventQueue.length} queued events`);

  const events = [...eventQueue];
  eventQueue = [];

  try {
    const response = await fetch(`${DOMINO_API_BASE}/excel-events/batch`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ events })
    });

    if (!response.ok) {
      // Re-queue on failure
      eventQueue = events.concat(eventQueue);
    }
  } catch (error) {
    // Re-queue on failure
    eventQueue = events.concat(eventQueue);
    isOnline = false;
  }
}

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
              console.log('✅ Model registered successfully');

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
 * Ribbon command: Mark selected cells as Input
 */
async function markAsInput() {
  await markSelection('input', '#E3F2FD'); // Light blue
}

/**
 * Ribbon command: Mark selected cells as Output
 */
async function markAsOutput() {
  await markSelection('output', '#E8F5E9'); // Light green
}

/**
 * Mark selected range as input or output
 */
async function markSelection(type, color) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('address');
      await context.sync();

      // Add to monitored cells
      await addMonitoredCell(range.address, type);

      // Visual feedback
      range.format.fill.color = color;
      await context.sync();

      console.log(`✅ Marked ${range.address} as ${type}`);

      // Notify Domino
      streamToDomino('cell_marked', {
        modelId: modelConfig?.id || 'unknown',
        range: range.address,
        type,
        timestamp: new Date().toISOString()
      });
    });
  } catch (error) {
    console.error(`Failed to mark as ${type}:`, error);
  }
}

/**
 * Add cell to monitored list
 */
async function addMonitoredCell(address, type) {
  // Initialize if needed
  if (!modelConfig) {
    modelConfig = { monitoredCells: [] };
  }

  if (!modelConfig.monitoredCells) {
    modelConfig.monitoredCells = [];
  }

  // Add to local config
  modelConfig.monitoredCells.push({
    range: address,
    type,
    addedAt: new Date().toISOString()
  });

  // Persist to Domino
  try {
    await fetch(`${DOMINO_API_BASE}/models/${modelConfig.id}/cells`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        range: address,
        type
      })
    });
  } catch (error) {
    console.error('Failed to persist monitored cell:', error);
  }
}

/**
 * Register ribbon command handlers
 */
Office.actions.associate("showRegisterModal", showRegisterModal);
Office.actions.associate("markAsInput", markAsInput);
Office.actions.associate("markAsOutput", markAsOutput);

console.log('📋 Command handlers registered');
