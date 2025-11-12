/**
 * Domino API Client
 * Handles all communication with Domino backend
 */

// Configuration - UPDATE THESE
const DOMINO_API_BASE = process.env.VITE_DOMINO_API_URL || 'https://your-domino.com/api';
const API_TIMEOUT = 10000; // 10 seconds

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

/**
 * Register a new model with Domino
 */
export async function registerModel(modelData) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        modelId: modelData.modelId,
        name: modelData.name,
        owner: modelData.owner,
        description: modelData.description || '',
        registeredAt: new Date().toISOString(),
        monitoredCells: []
      })
    });

    if (!response.ok) {
      throw new Error(`Registration failed: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to register model:', error);
    throw error;
  }
}

/**
 * Check if model is registered
 */
export async function getModelById(modelId) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models/${modelId}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      }
    });

    if (response.status === 404) {
      return null;
    }

    if (!response.ok) {
      throw new Error(`Failed to fetch model: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to fetch model:', error);
    return null;
  }
}

/**
 * Update model configuration
 */
export async function updateModel(modelId, updates) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models/${modelId}`, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(updates)
    });

    if (!response.ok) {
      throw new Error(`Update failed: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to update model:', error);
    throw error;
  }
}

/**
 * Add monitored cell to model
 */
export async function addMonitoredCell(modelId, cellData) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models/${modelId}/cells`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        range: cellData.range,
        type: cellData.type, // 'input' or 'output'
        addedAt: new Date().toISOString()
      })
    });

    if (!response.ok) {
      throw new Error(`Failed to add cell: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to add monitored cell:', error);
    throw error;
  }
}

/**
 * Remove monitored cell
 */
export async function removeMonitoredCell(modelId, cellRange) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models/${modelId}/cells/${encodeURIComponent(cellRange)}`, {
      method: 'DELETE',
      headers: {
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to remove cell: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to remove monitored cell:', error);
    throw error;
  }
}

/**
 * Stream event to Domino
 */
export async function streamEvent(eventData) {
  try {
    const response = await fetch(`${DOMINO_API_BASE}/excel-events`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(eventData),
      keepalive: true
    });

    return response.ok;
  } catch (error) {
    console.error('Failed to stream event:', error);
    return false;
  }
}

/**
 * Batch stream events (for offline queue flush)
 */
export async function streamEventBatch(events) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/excel-events/batch`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ events })
    });

    if (!response.ok) {
      throw new Error(`Batch stream failed: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to stream event batch:', error);
    throw error;
  }
}

/**
 * Get all registered models (for admin view)
 */
export async function getAllModels() {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models`, {
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
    console.error('Failed to fetch models:', error);
    return [];
  }
}

/**
 * Get model activity/events
 */
export async function getModelActivity(modelId, limit = 50) {
  try {
    const response = await fetchWithTimeout(`${DOMINO_API_BASE}/models/${modelId}/activity?limit=${limit}`, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch activity: ${response.statusText}`);
    }

    return await response.json();
  } catch (error) {
    console.error('Failed to fetch activity:', error);
    return [];
  }
}
