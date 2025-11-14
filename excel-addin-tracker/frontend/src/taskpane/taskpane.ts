/* global console, document, Excel, Office */

interface TrackedRange {
  name: string;
  range: string;
}

interface WorkbookModel {
  model_name: string;
  tracked_ranges: TrackedRange[];
  model_id: string;
  version: number;
}

const API_BASE_URL = "http://localhost:8000";
let currentModel: WorkbookModel | null = null;
let trackedRanges: TrackedRange[] = [];
let changeHandlers: Map<string, Excel.EventHandlerResult<Excel.WorksheetChangedEventArgs>> = new Map();

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("registerBtn")!.onclick = registerModel;
    document.getElementById("addRangeBtn")!.onclick = addTrackedRange;
    document.getElementById("updateModelBtn")!.onclick = updateModel;
    document.getElementById("loadModelBtn")!.onclick = loadModelFromWorkbook;
    document.getElementById("refreshBtn")!.onclick = refreshDisplay;

    // Load model on startup
    loadModelFromWorkbook();
  }
});

async function registerModel() {
  const modelName = (document.getElementById("modelName") as HTMLInputElement).value;

  if (!modelName) {
    showStatus("Please enter a model name", "error");
    return;
  }

  try {
    const response = await fetch(`${API_BASE_URL}/wb/upsert-model`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model_name: modelName,
        tracked_ranges: trackedRanges
      })
    });

    if (!response.ok) throw new Error("Failed to register model");

    const model: WorkbookModel = await response.json();
    currentModel = model;

    // Save model metadata to workbook
    await saveModelToWorkbook(model);

    showStatus(`Model registered successfully! ID: ${model.model_id}`, "success");
    refreshDisplay();

    // Set up change tracking
    await setupChangeTracking();
  } catch (error) {
    showStatus(`Error: ${error.message}`, "error");
  }
}

async function updateModel() {
  if (!currentModel) {
    showStatus("No model loaded", "error");
    return;
  }

  const modelName = (document.getElementById("modelName") as HTMLInputElement).value;

  try {
    const response = await fetch(`${API_BASE_URL}/wb/upsert-model`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model_name: modelName || currentModel.model_name,
        tracked_ranges: trackedRanges,
        model_id: currentModel.model_id,
        version: currentModel.version
      })
    });

    if (!response.ok) throw new Error("Failed to update model");

    const model: WorkbookModel = await response.json();
    currentModel = model;

    await saveModelToWorkbook(model);

    showStatus(`Model updated! New version: ${model.version}`, "success");
    refreshDisplay();

    // Re-setup change tracking with new ranges
    await setupChangeTracking();
  } catch (error) {
    showStatus(`Error: ${error.message}`, "error");
  }
}

async function loadModelFromWorkbook() {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const properties = workbook.properties.custom;
      properties.load("items");
      await context.sync();

      const modelIdProp = properties.items.find(p => p.key === "ModelId");

      if (!modelIdProp) {
        showStatus("No model found in this workbook", "error");
        return;
      }

      const modelId = modelIdProp.value;

      // Load from backend
      const response = await fetch(`${API_BASE_URL}/wb/load-model?model_id=${modelId}`);

      if (!response.ok) throw new Error("Failed to load model");

      const model: WorkbookModel = await response.json();
      currentModel = model;
      trackedRanges = [...model.tracked_ranges];

      showStatus("Model loaded successfully", "success");
      refreshDisplay();

      // Set up change tracking
      await setupChangeTracking();
    });
  } catch (error) {
    showStatus(`Error loading model: ${error.message}`, "error");
  }
}

async function saveModelToWorkbook(model: WorkbookModel) {
  await Excel.run(async (context) => {
    const workbook = context.workbook;
    const properties = workbook.properties.custom;

    // Remove existing properties
    properties.load("items");
    await context.sync();

    const existingProps = properties.items.filter(p =>
      p.key === "ModelId" || p.key === "ModelName" || p.key === "ModelVersion"
    );
    existingProps.forEach(p => p.delete());
    await context.sync();

    // Add new properties
    properties.add("ModelId", model.model_id);
    properties.add("ModelName", model.model_name);
    properties.add("ModelVersion", model.version.toString());

    await context.sync();
  });
}

function addTrackedRange() {
  const name = (document.getElementById("rangeName") as HTMLInputElement).value;
  const range = (document.getElementById("rangeAddress") as HTMLInputElement).value;

  if (!name || !range) {
    showStatus("Please enter both range name and address", "error");
    return;
  }

  trackedRanges.push({ name, range });

  (document.getElementById("rangeName") as HTMLInputElement).value = "";
  (document.getElementById("rangeAddress") as HTMLInputElement).value = "";

  renderTrackedRanges();
  showStatus("Tracked range added", "success");
}

function removeTrackedRange(index: number) {
  trackedRanges.splice(index, 1);
  renderTrackedRanges();
  showStatus("Tracked range removed", "success");
}

function renderTrackedRanges() {
  const container = document.getElementById("trackedRangesList")!;

  if (trackedRanges.length === 0) {
    container.innerHTML = "<p style='color: #666;'>No tracked ranges yet</p>";
    return;
  }

  container.innerHTML = trackedRanges.map((tr, index) => `
    <div class="tracked-range">
      <div class="tracked-range-info">
        <div class="tracked-range-name">${tr.name}</div>
        <div class="tracked-range-address">${tr.range}</div>
      </div>
      <button class="remove-btn" onclick="removeTrackedRange(${index})">Remove</button>
    </div>
  `).join("");
}

async function setupChangeTracking() {
  if (!currentModel) return;

  await Excel.run(async (context) => {
    // Clear existing handlers
    for (const [sheetName, handler] of changeHandlers.entries()) {
      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();
      if (!sheet.isNullObject) {
        sheet.onChanged.remove(handler);
      }
    }
    changeHandlers.clear();

    // Set up new handlers for each tracked range
    for (const trackedRange of currentModel.tracked_ranges) {
      const [sheetName] = trackedRange.range.split("!");

      const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
      await context.sync();

      if (!sheet.isNullObject) {
        const handler = sheet.onChanged.add(async (event) => {
          await handleRangeChange(event, trackedRange);
        });
        changeHandlers.set(sheetName, handler);
      }
    }

    await context.sync();
  });
}

async function handleRangeChange(
  event: Excel.WorksheetChangedEventArgs,
  trackedRange: TrackedRange
) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem(event.worksheetId).getRange(event.address);
      range.load("values, address");
      await context.sync();

      // Check if changed range intersects with tracked range
      const [sheetName, rangeAddress] = trackedRange.range.split("!");
      const trackedRangeObj = context.workbook.worksheets.getItem(event.worksheetId).getRange(rangeAddress);
      const intersection = trackedRangeObj.getIntersectionOrNullObject(range);
      intersection.load("address");
      await context.sync();

      if (!intersection.isNullObject) {
        // Create trace
        await createTrace(trackedRange.name, range.values);
      }
    });
  } catch (error) {
    console.error("Error handling range change:", error);
  }
}

async function createTrace(trackedRangeName: string, value: any) {
  if (!currentModel) return;

  try {
    const response = await fetch(`${API_BASE_URL}/wb/create-model-trace`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model_id: currentModel.model_id,
        timestamp: new Date().toISOString(),
        tracked_range_name: trackedRangeName,
        username: "User",  // In production, get from Office context
        value: value
      })
    });

    if (!response.ok) {
      console.error("Failed to create trace");
    }
  } catch (error) {
    console.error("Error creating trace:", error);
  }
}

function refreshDisplay() {
  if (currentModel) {
    document.getElementById("modelInfo")!.classList.remove("hidden");
    document.getElementById("currentModelName")!.textContent = currentModel.model_name;
    document.getElementById("currentModelId")!.textContent = currentModel.model_id;
    document.getElementById("currentVersion")!.textContent = currentModel.version.toString();
    document.getElementById("updateModelBtn")!.classList.remove("hidden");
    (document.getElementById("modelName") as HTMLInputElement).value = currentModel.model_name;
  } else {
    document.getElementById("modelInfo")!.classList.add("hidden");
    document.getElementById("updateModelBtn")!.classList.add("hidden");
  }

  renderTrackedRanges();
}

function showStatus(message: string, type: "success" | "error") {
  const statusEl = document.getElementById("status")!;
  statusEl.textContent = message;
  statusEl.className = type;
  statusEl.style.display = "block";

  setTimeout(() => {
    statusEl.style.display = "none";
  }, 5000);
}

// Make functions globally accessible
(window as any).removeTrackedRange = removeTrackedRange;
