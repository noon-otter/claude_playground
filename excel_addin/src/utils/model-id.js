/**
 * Model ID Management
 * Handles persistent model identification across Excel file versions
 */

const MODEL_ID_PROPERTY = 'DominoModelId';
const MODEL_NAME_PROPERTY = 'DominoModelName';

/**
 * Get or create persistent model ID
 * Stored in Excel's custom document properties - survives Save As, renames, etc.
 */
export async function getOrCreateModelId() {
  try {
    // Check if Excel APIs are available
    if (typeof Excel === 'undefined' || !Excel.run) {
      console.warn('[model-id] Excel APIs not available, using fallback');
      return generateModelId(); // Fallback to generating new ID
    }

    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const properties = workbook.properties.custom;

      properties.load('items');
      await context.sync();

    // Check if model ID exists
    let modelIdProp = null;
    for (let i = 0; i < properties.items.length; i++) {
      if (properties.items[i].key === MODEL_ID_PROPERTY) {
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
      properties.add(MODEL_ID_PROPERTY, modelId);
      await context.sync();

      return modelId;
    });
  } catch (error) {
    console.error('[model-id] Error accessing Excel properties:', error);
    console.log('[model-id] Falling back to session-based ID');
    // Fallback: use localStorage or generate temporary ID
    let fallbackId = localStorage.getItem('DominoModelId_Fallback');
    if (!fallbackId) {
      fallbackId = generateModelId();
      localStorage.setItem('DominoModelId_Fallback', fallbackId);
    }
    return fallbackId;
  }
}

/**
 * Generate unique model ID
 * Format: excel_<timestamp>_<random>
 */
function generateModelId() {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substring(2, 9);
  return `excel_${timestamp}_${random}`;
}

/**
 * Get model ID if it exists (doesn't create new one)
 */
export async function getModelId() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const properties = workbook.properties.custom;

      properties.load('items');
      await context.sync();

      for (let i = 0; i < properties.items.length; i++) {
        if (properties.items[i].key === MODEL_ID_PROPERTY) {
          const prop = properties.items[i];
          prop.load('value');
          await context.sync();
          return prop.value;
        }
      }

      return null;
    });
  } catch (error) {
    console.error('Failed to get model ID:', error);
    return null;
  }
}

/**
 * Set model name in document properties
 */
export async function setModelName(name) {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const properties = workbook.properties.custom;

    properties.load('items');
    await context.sync();

    // Check if property exists
    let nameProp = null;
    for (let i = 0; i < properties.items.length; i++) {
      if (properties.items[i].key === MODEL_NAME_PROPERTY) {
        nameProp = properties.items[i];
        break;
      }
    }

    if (nameProp) {
      // Update existing
      nameProp.delete();
      await context.sync();
    }

    // Add new value
    properties.add(MODEL_NAME_PROPERTY, name);
    await context.sync();

    return name;
  });
}

/**
 * Get model name from document properties
 */
export async function getModelName() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const properties = workbook.properties.custom;

      properties.load('items');
      await context.sync();

      for (let i = 0; i < properties.items.length; i++) {
        if (properties.items[i].key === MODEL_NAME_PROPERTY) {
          const prop = properties.items[i];
          prop.load('value');
          await context.sync();
          return prop.value;
        }
      }

      return null;
    });
  } catch (error) {
    console.error('Failed to get model name:', error);
    return null;
  }
}

/**
 * Get workbook file name
 */
export async function getWorkbookName() {
  try {
    if (typeof Excel === 'undefined' || !Excel.run) {
      return 'Excel Workbook';
    }

    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load('name');
      await context.sync();
      return workbook.name;
    });
  } catch (error) {
    console.error('Failed to get workbook name:', error);
    return 'Excel Workbook';
  }
}

/**
 * Get file path (if available)
 */
export function getFilePath() {
  try {
    return Office.context.document.url || null;
  } catch {
    return null;
  }
}

/**
 * Get all custom properties (for debugging)
 */
export async function getAllCustomProperties() {
  try {
    return await Excel.run(async (context) => {
      const workbook = context.workbook;
      const properties = workbook.properties.custom;

      properties.load('items');
      await context.sync();

      const props = {};
      for (let i = 0; i < properties.items.length; i++) {
        const prop = properties.items[i];
        prop.load(['key', 'value']);
        await context.sync();
        props[prop.key] = prop.value;
      }

      return props;
    });
  } catch (error) {
    console.error('Failed to get custom properties:', error);
    return {};
  }
}
