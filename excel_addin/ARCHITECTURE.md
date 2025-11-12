# Architecture Documentation

Technical architecture of the Domino Excel Governance Add-in.

## Overview

The add-in uses a **dual-runtime architecture**:
1. **Background runtime** - Always running, independent of UI
2. **Taskpane runtime** - React UI, optional for users

This enables always-on monitoring without requiring the taskpane to be open.

## System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         Excel Client                            │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │  User Interactions (cell edits, selections, saves)       │  │
│  └──────────────────┬───────────────────────────────────────┘  │
│                     │                                           │
│                     ▼                                           │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │         Office.js Event System                           │  │
│  │  (onChanged, onSelectionChanged, onAutoSave, etc.)       │  │
│  └──────────────────┬───────────────────────────────────────┘  │
│                     │                                           │
│         ┌───────────┴──────────┐                                │
│         │                      │                                │
│         ▼                      ▼                                │
│  ┌──────────────┐      ┌──────────────┐                        │
│  │  commands.js │      │ Taskpane UI  │                        │
│  │  (Background)│      │   (React)    │                        │
│  │              │      │              │                        │
│  │ - Monitor    │      │ - Dashboard  │                        │
│  │ - Stream     │      │ - Register   │                        │
│  │ - Queue      │      │ - View cells │                        │
│  └──────┬───────┘      └──────┬───────┘                        │
│         │                     │                                │
│         └─────────┬───────────┘                                │
│                   │                                             │
└───────────────────┼─────────────────────────────────────────────┘
                    │
                    ▼ HTTP POST (fetch)
         ┌──────────────────────┐
         │    Domino API        │
         │  (FastAPI/Flask)     │
         │                      │
         │ - Store events       │
         │ - Track models       │
         │ - Compliance checks  │
         └──────────┬───────────┘
                    │
                    ▼
         ┌──────────────────────┐
         │     Database         │
         │  (Postgres/MongoDB)  │
         └──────────────────────┘
```

## Component Breakdown

### 1. Manifest (manifest.xml)

The manifest defines the add-in structure and capabilities.

**Key sections:**
```xml
<Runtimes>
  <Runtime lifetime="long">
    <!-- Background script runs continuously -->
  </Runtime>
</Runtimes>

<ExtensionPoint xsi:type="PrimaryCommandSurface">
  <!-- Ribbon buttons -->
</ExtensionPoint>
```

**Runtime types:**
- `lifetime="long"` - Persistent background runtime
- `lifetime="short"` - UI-bound runtime (taskpane)

### 2. Background Script (commands.js)

**Purpose:** Always-on monitoring independent of UI.

**Execution flow:**
```
Excel loads
    ↓
Office.onReady() fires
    ↓
initializeMonitoring()
    ↓
getOrCreateModelId() - Check custom properties
    ↓
checkModelRegistration() - Query Domino API
    ↓
If registered:
    startLiveMonitoring() - Attach event handlers
    ↓
    Events stream continuously to Domino
```

**Event handlers registered:**
- `workbook.worksheets.onChanged` → Cell changes
- `workbook.onSelectionChanged` → User interactions
- `workbook.onAutoSaveSettingChanged` → Saves
- `worksheets.onActivated` → Worksheet switches

**Key design decisions:**

1. **No blocking operations** - All API calls use `fetch()` with no await in event handlers
2. **Fire-and-forget streaming** - Events sent without waiting for response
3. **Offline queue** - Failed requests queued locally (max 100 events)
4. **Keepalive flag** - Ensures events sent even during page unload

### 3. Persistent Model ID (model-id.js)

**Problem:** Excel files can be:
- Saved As (new filename)
- Renamed
- Copied
- Synced across devices

**Solution:** Custom Document Properties

```javascript
workbook.properties.custom.add('DominoModelId', 'excel_abc123')
```

**Properties:**
- Stored inside the .xlsx file itself
- Survives all file operations
- Accessible via Office.js API
- Invisible to normal users

**ID format:**
```
excel_<timestamp_base36>_<random>
Example: excel_l5k3m9_ab7cd2e
```

Provides:
- Uniqueness (timestamp + random)
- Readability (base36, short)
- Namespace (prefix 'excel_')

### 4. React UI Components

**Component tree:**
```
App.jsx
  ├─ Status badge
  ├─ Model info card
  └─ MonitorView.jsx
       ├─ Model details
       ├─ Monitored cells list
       └─ Recent activity feed

RegisterModal.jsx
  ├─ Model name input
  ├─ Owner email input
  ├─ Description textarea
  └─ Submit → API call
```

**State management:**
- No Redux/Context needed (simple app)
- `useState` for local state
- Props for parent-child communication
- API calls in components directly

**UI Library:** Fluent UI (Microsoft's design system)
- Native Office look and feel
- Accessible components
- Consistent with Excel UI

### 5. API Client (domino-api.js)

**Responsibilities:**
- All HTTP communication with Domino
- Timeout handling (10s default)
- Error recovery
- Type safety (Pydantic-like validation via duck typing)

**Key functions:**
```javascript
registerModel(data)        // POST /api/models
getModelById(id)           // GET /api/models/:id
streamEvent(event)         // POST /api/excel-events
streamEventBatch(events)   // POST /api/excel-events/batch
addMonitoredCell(id, cell) // POST /api/models/:id/cells
```

**Error handling:**
- Timeout after 10s
- Return `null` on 404 (not an error)
- Throw on 4xx/5xx
- Log but don't crash on network errors

### 6. Event Streaming

**Design principle:** Optimistic, non-blocking

```javascript
function streamToDomino(eventType, data) {
  fetch(DOMINO_API, {
    method: 'POST',
    body: JSON.stringify({ event: eventType, ...data }),
    keepalive: true
  })
  .catch(err => queueEventLocally(eventType, data));
  // No await - fire and forget
}
```

**Keepalive flag:**
- Tells browser to complete request even if page closes
- Critical for save/close events
- Supported in modern browsers

**Offline queue:**
```
Event fails
    ↓
Add to local array
    ↓
Retry in 30s
    ↓
If successful: Batch flush all queued events
```

**Queue limits:**
- Max 100 events stored
- FIFO eviction (oldest dropped first)
- Automatic flush on reconnect

## Data Flow

### Registration Flow

```
User clicks "Register Model"
    ↓
Office.context.ui.displayDialogAsync()
    ↓
register.html loads in modal
    ↓
RegisterModal.jsx renders
    ↓
User fills form → Submit
    ↓
POST /api/models
    ↓
Success: messageParent({ action: 'registered', config })
    ↓
commands.js receives message
    ↓
Reload config, start monitoring
```

### Cell Change Flow

```
User edits cell A1: 100 → 200
    ↓
Excel fires onChanged event
    ↓
handleCellChange(event)
    ↓
Check if A1 in monitoredCells
    ↓
If yes:
    getCellValue(A1) → 200
    ↓
    streamToDomino('cell_changed', {
      cell: 'A1',
      value: 200,
      type: 'input'
    })
    ↓
POST /api/excel-events
    ↓
Domino stores event
```

### Offline → Online Flow

```
Network disconnects
    ↓
streamToDomino() fails
    ↓
queueEventLocally() adds to array
    ↓
[Multiple events queued...]
    ↓
Network reconnects
    ↓
Next streamToDomino() succeeds
    ↓
flushEventQueue() triggered
    ↓
POST /api/excel-events/batch
    ↓
All queued events sent at once
```

## Security Model

### Client-side (Add-in)

**No authentication in MVP** - Relies on:
- Office 365 user context (email)
- Network-level security (VPN, internal network)
- CORS restrictions

**Future:** OAuth 2.0 with MSAL.js
```javascript
import { PublicClientApplication } from "@azure/msal-browser";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: "your-app-id",
    authority: "https://login.microsoftonline.com/your-tenant"
  }
});

// Get token
const token = await msalInstance.acquireTokenSilent({
  scopes: ["api://domino/.default"]
});

// Add to requests
fetch(DOMINO_API, {
  headers: {
    'Authorization': `Bearer ${token.accessToken}`
  }
});
```

### Server-side (Domino API)

**Current:** Open API (internal network only)

**Production:**
- JWT validation
- Role-based access control
- Rate limiting
- Input validation (Pydantic models)

## Performance Considerations

### Bundle Size

Current (estimated):
- React + ReactDOM: ~140 KB
- Fluent UI: ~200 KB
- Office.js: CDN (not bundled)
- Total: ~350 KB gzipped

**Optimizations:**
- Code splitting (separate bundles for taskpane, modal, commands)
- Tree shaking (Vite default)
- Lazy loading for dashboard

### Event Throttling

Some events fire rapidly (e.g., `onSelectionChanged` as user drags).

**Solution:** Debouncing
```javascript
let selectionTimeout;
workbook.onSelectionChanged.add((event) => {
  clearTimeout(selectionTimeout);
  selectionTimeout = setTimeout(() => {
    streamToDomino('selection_changed', event);
  }, 500); // Wait 500ms of no changes
});
```

### API Rate Limits

If Domino API has rate limits, implement client-side batching:
```javascript
let eventBuffer = [];
setInterval(() => {
  if (eventBuffer.length > 0) {
    streamEventBatch(eventBuffer);
    eventBuffer = [];
  }
}, 5000); // Batch every 5s
```

## Extensibility

### Adding New Event Types

1. Add handler in `commands.js`:
```javascript
workbook.onCalculated.add((event) => {
  streamToDomino('calculation_complete', {
    modelId,
    timestamp: new Date().toISOString()
  });
});
```

2. Handle in Domino API:
```python
@app.post("/api/excel-events")
def receive_event(event: ExcelEvent):
    if event.event == "calculation_complete":
        # Handle calculation event
        pass
```

### Adding New UI Components

1. Create component in `src/taskpane/`:
```javascript
// ComplianceReport.jsx
export default function ComplianceReport({ modelId }) {
  // Fetch compliance data
  // Render report
}
```

2. Import in `App.jsx`:
```javascript
import ComplianceReport from './ComplianceReport';

// Add to UI
<ComplianceReport modelId={modelId} />
```

### Adding Ribbon Buttons

1. Update `manifest.xml`:
```xml
<Control xsi:type="Button" id="NewButton">
  <Label resid="NewButton.Label"/>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>newButtonHandler</FunctionName>
  </Action>
</Control>
```

2. Register in `commands.js`:
```javascript
function newButtonHandler() {
  // Implementation
}

Office.actions.associate("newButtonHandler", newButtonHandler);
```

## Testing Strategy

### Unit Tests (Future)

```javascript
// vitest.config.js
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'jsdom',
    setupFiles: ['./tests/setup.js']
  }
});
```

Mock Office.js:
```javascript
// tests/setup.js
global.Office = {
  onReady: (cb) => cb({ host: 'Excel' }),
  context: {
    document: { url: 'test.xlsx' }
  }
};

global.Excel = {
  run: async (callback) => {
    const mockContext = {
      workbook: mockWorkbook,
      sync: async () => {}
    };
    return await callback(mockContext);
  }
};
```

### Integration Tests

Use Office.js Script Lab for live testing:
1. Install Script Lab add-in
2. Paste code snippets
3. Run in real Excel
4. Verify events fire correctly

### E2E Tests (Future)

Playwright with Office Online:
```javascript
test('register model', async ({ page }) => {
  await page.goto('https://office.com/launch/excel');
  await page.click('text=Domino');
  await page.click('text=Register Model');
  // ...
});
```

## Monitoring & Observability

### Client-side Logging

```javascript
// Add structured logging
function log(level, message, data) {
  console[level](message, data);

  // Send to telemetry
  fetch('/api/telemetry', {
    method: 'POST',
    body: JSON.stringify({
      level,
      message,
      data,
      timestamp: Date.now(),
      modelId: currentModelId
    })
  });
}
```

### Server-side Metrics

Track:
- Events per second
- Registration rate
- Error rate by type
- Active models
- Offline queue size (if retrievable)

## Future Enhancements

### Formula Dependency Tracking

```javascript
range.load('formulas');
await context.sync();

// Parse formula, extract dependencies
const deps = parseFormula(range.formulas[0][0]);
// Track: B1 depends on A1, A2
```

### Conflict Detection

```javascript
// Multiple users editing same cell
workbook.onChanged.add((event) => {
  const lastEdit = getLastEdit(event.address);
  if (lastEdit.user !== currentUser && Date.now() - lastEdit.time < 5000) {
    showWarning('Another user just edited this cell');
  }
});
```

### Rollback Capability

```javascript
// Store history
streamToDomino('cell_changed', {
  cell: 'A1',
  oldValue: 100,
  newValue: 200
});

// Rollback endpoint
async function rollbackCell(cell, toTimestamp) {
  const history = await getHistory(cell);
  const value = history.find(h => h.timestamp === toTimestamp).value;
  // Set cell value
}
```

## Limitations & Known Issues

1. **No cross-workbook tracking** - Each file monitored independently
2. **Limited in Excel Online** - Some APIs unavailable (e.g., VBA)
3. **No real-time collaboration sync** - Events don't propagate between co-editors
4. **Queue persistence** - Offline queue lost on Excel close
5. **No encryption** - Events sent in plain HTTPS (add E2E encryption for sensitive data)

## References

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API](https://docs.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Fluent UI React](https://react.fluentui.dev/)
- [Vite Documentation](https://vitejs.dev/)
