/**
 * Domino API Client - Architecture Compliant Version
 * Implements the exact API specification from DEPLOYMENT.md
 *
 * Endpoints:
 *   PUT  /wb/upsert-model         - Create or update model
 *   GET  /wb/load-model           - Load model metadata
 *   POST /wb/create-model-trace   - Create trace log entry
 */

// Configuration - UPDATE THESE
const DOMINO_API_BASE = import.meta.env.VITE_DOMINO_API_URL || 'http://localhost:5000';
const API_TIMEOUT = 10000; // 10 seconds

console.log('[domino-api-v2.js] API Base URL:', DOMINO_API_BASE);

/**
 * Fetch with timeout
 */
function fetchWithTimeout(url, options = {}, timeout = API_TIMEOUT) {
  return Promise.race([
    fetch(url, options),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error('Request timeout')), timeout)
    )
  ]);
}

// =====================================================
// CORE API ENDPOINTS (Per Architecture)
// =====================================================

/**
 * PUT /wb/upsert-model
 *
 * Used for:
 * - Creating a model
 * - Updating a model's name/tracked ranges
 * - Returning versioned model metadata
 *
 * @param {Object} data
 * @param {string} data.model_name - Model name
 * @param {Array<{name: string, range: string}>} data.tracked_ranges - Tracked ranges
 * @param {string} [data.model_id] - Optional: existing model ID (for update)
 * @param {number} [data.version] - Optional: current version (for update)
 *
 * @returns {Promise<{model_name, tracked_ranges, model_id, version}>}
 */
export async function upsertModel(data) {
  try {
    console.log('[domino-api] PUT /wb/upsert-model', data);

    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/wb/upsert-model`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      throw new Error(`Upsert failed: ${response.statusText}`);
    }

    const result = await response.json();
    console.log('[domino-api] Upsert result:', result);

    return result;
  } catch (error) {
    console.error('[domino-api] Failed to upsert model:', error);
    throw error;
  }
}

/**
 * GET /wb/load-model
 *
 * Used for:
 * - Loading metadata for an existing model
 *
 * @param {string} modelId - Model ID to load
 * @returns {Promise<{model_name, tracked_ranges, model_id, version} | null>}
 */
export async function loadModel(modelId) {
  try {
    console.log('[domino-api] GET /wb/load-model?model_id=', modelId);

    const response = await fetchWithTimeout(
      `${DOMINO_API_BASE}/wb/load-model?model_id=${encodeURIComponent(modelId)}`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.status === 404) {
      console.log('[domino-api] Model not found:', modelId);
      return null;
    }

    if (!response.ok) {
      throw new Error(`Load failed: ${response.statusText}`);
    }

    const result = await response.json();
    console.log('[domino-api] Load result:', result);

    return result;
  } catch (error) {
    console.error('[domino-api] Failed to load model:', error);
    return null;
  }
}

/**
 * POST /wb/create-model-trace
 *
 * Triggered when a tracked range changes.
 *
 * @param {Object} data
 * @param {string} data.model_id - Model ID
 * @param {string} data.timestamp - ISO timestamp
 * @param {string} data.tracked_range_name - Name of tracked range
 * @param {string} data.username - Username
 * @param {any} data.value - Cell value
 *
 * @returns {Promise<{success: boolean}>}
 */
export async function createModelTrace(data) {
  try {
    console.log('[domino-api] POST /wb/create-model-trace', data);

    const response = await fetch(`${DOMINO_API_BASE}/wb/create-model-trace`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(data),
      keepalive: true // Survives page navigation
    });

    if (!response.ok) {
      throw new Error(`Trace creation failed: ${response.statusText}`);
    }

    const result = await response.json();
    return result;
  } catch (error) {
    console.error('[domino-api] Failed to create trace:', error);
    return { success: false };
  }
}

/**
 * POST /wb/create-model-trace-batch
 *
 * Batch version for multiple tracked range changes at once.
 *
 * @param {Object} data
 * @param {string} data.model_id - Model ID
 * @param {string} data.timestamp - ISO timestamp
 * @param {Array<{tracked_range_name: string, value: any}>} data.changes - Array of changes
 * @param {string} data.username - Username
 *
 * @returns {Promise<{success: boolean}>}
 */
export async function createModelTraceBatch(data) {
  try {
    console.log('[domino-api] POST /wb/create-model-trace-batch', data);

    const response = await fetchWithTimeout(
      `${DOMINO_API_BASE}/wb/create-model-trace-batch`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
      }
    );

    if (!response.ok) {
      throw new Error(`Batch trace creation failed: ${response.statusText}`);
    }

    const result = await response.json();
    return result;
  } catch (error) {
    console.error('[domino-api] Failed to create batch trace:', error);
    return { success: false };
  }
}

// =====================================================
// UTILITY FUNCTIONS (For debugging/monitoring)
// =====================================================

/**
 * Get all registered models (for debugging)
 */
export async function getAllModels() {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/wb/models`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch models: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('[domino-api] Failed to fetch models:', error);
    return [];
  }
}

/**
 * Get model traces (for debugging)
 */
export async function getModelTraces(modelId, limit = 50) {
  try {
    const response = await fetchWithTimeout(
      `${DOMINO_API_BASE}/wb/traces/${encodeURIComponent(modelId)}?limit=${limit}`,
      {
        method: 'GET',
        headers: {
          'Content-Type': 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch traces: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('[domino-api] Failed to fetch traces:', error);
    return [];
  }
}

// =====================================================
// BACKWARDS COMPATIBILITY LAYER
// (Optional: Remove once all code is migrated)
// =====================================================

/**
 * @deprecated Use upsertModel() instead
 */
export async function registerModel(modelData) {
  console.warn('[domino-api] registerModel() is deprecated, use upsertModel()');

  return await upsertModel({
    model_name: modelData.name,
    tracked_ranges: [], // Start with empty tracked ranges
    model_id: modelData.modelId
  });
}

/**
 * @deprecated Use loadModel() instead
 */
export async function getModelById(modelId) {
  console.warn('[domino-api] getModelById() is deprecated, use loadModel()');
  return await loadModel(modelId);
}

/**
 * @deprecated Use upsertModel() with model_id and version instead
 */
export async function updateModel(modelId, updates) {
  console.warn('[domino-api] updateModel() is deprecated, use upsertModel()');

  // Load current model first
  const currentModel = await loadModel(modelId);
  if (!currentModel) {
    throw new Error(`Model not found: ${modelId}`);
  }

  // Merge updates
  return await upsertModel({
    model_name: updates.model_name || currentModel.model_name,
    tracked_ranges: updates.tracked_ranges || currentModel.tracked_ranges,
    model_id: modelId,
    version: currentModel.version
  });
}

/**
 * @deprecated Use upsertModel() to add tracked ranges instead
 */
export async function addMonitoredCell(modelId, cellData) {
  console.warn('[domino-api] addMonitoredCell() is deprecated, use upsertModel()');

  // Load current model
  const currentModel = await loadModel(modelId);
  if (!currentModel) {
    throw new Error(`Model not found: ${modelId}`);
  }

  // Add new tracked range
  const newTrackedRange = {
    name: cellData.range, // Use range as name if no name provided
    range: cellData.range
  };

  return await upsertModel({
    model_name: currentModel.model_name,
    tracked_ranges: [...currentModel.tracked_ranges, newTrackedRange],
    model_id: modelId,
    version: currentModel.version
  });
}

/**
 * @deprecated Use upsertModel() to remove tracked ranges instead
 */
export async function removeMonitoredCell(modelId, cellRange) {
  console.warn('[domino-api] removeMonitoredCell() is deprecated, use upsertModel()');

  // Load current model
  const currentModel = await loadModel(modelId);
  if (!currentModel) {
    throw new Error(`Model not found: ${modelId}`);
  }

  // Remove tracked range
  const updatedRanges = currentModel.tracked_ranges.filter(
    tr => tr.range !== cellRange
  );

  return await upsertModel({
    model_name: currentModel.model_name,
    tracked_ranges: updatedRanges,
    model_id: modelId,
    version: currentModel.version
  });
}
