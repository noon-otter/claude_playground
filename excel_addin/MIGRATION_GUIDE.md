# Migration Guide: Architecture-Compliant Implementation

## Overview

This guide explains the changes made to align your Excel Add-in with the architecture specification in `DEPLOYMENT.md`.

## Summary of Changes

### 1. **New Files Created**

#### TypeScript Type Definitions
- **File**: `src/types/model.ts`
- **Purpose**: Complete type definitions matching the architecture
- **Exports**:
  - `TrackedRange`: `{name: string, range: string}`
  - `WorkbookModel`: `{model_name, tracked_ranges[], model_id, version}`
  - `WorkbookTrace`: `{model_id, timestamp, tracked_range_name, username, value}`
  - Request/Response types for all API endpoints

#### Database Schema
- **File**: `database_schema.sql`
- **Purpose**: SQL DDL for database tables
- **Tables**:
  - `dbo.workbook_model`: Model metadata with versioning
  - `dbo.workbook_trace`: Trace log entries
- **Features**: Indexes, foreign keys, auto-update triggers

#### Architecture-Compliant Backend
- **File**: `domino-api-backend.py`
- **Purpose**: FastAPI implementation matching exact API spec
- **Endpoints**:
  - `PUT /wb/upsert-model` - Create/update models with versioning
  - `GET /wb/load-model` - Load model metadata
  - `POST /wb/create-model-trace` - Create trace entries
  - `POST /wb/create-model-trace-batch` - Batch trace creation

#### New API Client
- **File**: `src/utils/domino-api-v2.js`
- **Purpose**: Frontend API client using correct endpoints
- **Functions**:
  - `upsertModel()` - Create/update models
  - `loadModel()` - Load model by ID
  - `createModelTrace()` - Create single trace
  - `createModelTraceBatch()` - Create multiple traces
- **Includes**: Backwards compatibility layer for old code

#### New Commands Handler
- **File**: `src/commands/commands-v2.js`
- **Purpose**: Background monitoring using correct architecture
- **Event Flows**:
  - Workbook Load → `GET /wb/load-model`
  - Register Model → `PUT /wb/upsert-model` (version=1)
  - Add Tracked Range → `PUT /wb/upsert-model` (increment version)
  - Range Change → `POST /wb/create-model-trace`

### 2. **Updated Files**

#### Frontend Components
- **`src/taskpane/App.jsx`**:
  - Changed import: `loadModel` from `domino-api-v2`
  - Uses new model structure: `model_name`, `tracked_ranges`, `version`

- **`src/taskpane/RegisterModal.jsx`**:
  - Changed import: `upsertModel` from `domino-api-v2`
  - Sends: `{model_name, tracked_ranges: [], model_id}`

- **`src/taskpane/MonitorView.jsx`**:
  - Uses `tracked_ranges` instead of `monitoredCells`
  - Displays: `{name, range}` instead of `{range, type, addedAt}`
  - Shows `version` field in Model Info
  - Loads traces from `getModelTraces()`

## Key Architectural Changes

### Data Structure Changes

| Old (monitoredCells) | New (tracked_ranges) |
|----------------------|----------------------|
| `{range, type, addedAt}` | `{name, range}` |
| Type: 'input' or 'output' | No type field |
| Tracked by cell type | Tracked by named range |

### Model Fields

| Old | New |
|-----|-----|
| `modelId` | `model_id` |
| `name` | `model_name` |
| `owner` | (removed) |
| `description` | (removed) |
| `registeredAt` | (removed) |
| `monitoredCells[]` | `tracked_ranges[]` |
| (none) | `version` (new!) |

### API Endpoints

| Old | New | Purpose |
|-----|-----|---------|
| `POST /api/models` | `PUT /wb/upsert-model` | Create model |
| `GET /api/models/{id}` | `GET /wb/load-model?model_id={id}` | Load model |
| `PATCH /api/models/{id}` | `PUT /wb/upsert-model` | Update model |
| `POST /api/models/{id}/cells` | `PUT /wb/upsert-model` | Add tracked range |
| `DELETE /api/models/{id}/cells/{range}` | `PUT /wb/upsert-model` | Remove tracked range |
| `POST /api/excel-events` | `POST /wb/create-model-trace` | Log trace |

### Event Flow Changes

#### Before (Old Architecture)
```
Cell Change → POST /api/excel-events
{
  event: "cell_changed",
  modelId: "...",
  cell: "A1",
  value: 100,
  type: "input",
  user: "user@example.com"
}
```

#### After (New Architecture)
```
Tracked Range Change → POST /wb/create-model-trace
{
  model_id: "excel_123_abc",
  timestamp: "2025-01-15T10:30:00Z",
  tracked_range_name: "Revenue",
  username: "user@example.com",
  value: 100
}
```

## Migration Steps

### Step 1: Database Setup
```sql
-- Run database_schema.sql to create tables
-- This creates:
-- - dbo.workbook_model
-- - dbo.workbook_trace
```

### Step 2: Start New Backend
```bash
cd excel_addin
pip install fastapi uvicorn
python domino-api-backend.py
# Backend runs on http://localhost:5000
```

### Step 3: Update Frontend Configuration

**Option A: Keep both versions (recommended for testing)**
- Keep old files: `domino-api.js`, `commands.js`
- Use new files: `domino-api-v2.js`, `commands-v2.js`
- Update imports in components to use `-v2` versions

**Option B: Replace old files**
- Replace `src/utils/domino-api.js` with `domino-api-v2.js`
- Replace `src/commands/commands.js` with `commands-v2.js`
- Update all imports to remove `-v2` suffix

### Step 4: Update Ribbon Commands

Edit `manifest.xml` to update ribbon button actions:

**Old:**
```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>markAsInput</FunctionName>
</Action>
<Action xsi:type="ExecuteFunction">
  <FunctionName>markAsOutput</FunctionName>
</Action>
```

**New:**
```xml
<Action xsi:type="ExecuteFunction">
  <FunctionName>addTrackedRange</FunctionName>
</Action>
```

### Step 5: Test Migration

1. **Test Model Registration**
   - Open Excel workbook
   - Click "Register Model"
   - Verify model created with `version: 1`

2. **Test Load Model**
   - Close and reopen workbook
   - Verify model loads with tracked ranges

3. **Test Add Tracked Range**
   - Select cell range
   - Click "Add Tracked Range"
   - Enter range name
   - Verify model version increments

4. **Test Trace Logging**
   - Change value in tracked range
   - Verify trace created in backend logs

## Backwards Compatibility

The new `domino-api-v2.js` includes a backwards compatibility layer:

```javascript
// Old code still works (with deprecation warnings)
await registerModel({...})    // → upsertModel()
await getModelById(id)         // → loadModel()
await updateModel(id, {...})   // → upsertModel()
await addMonitoredCell(...)    // → upsertModel()
await removeMonitoredCell(...) // → upsertModel()
```

## Breaking Changes

### 1. Ribbon Commands
- **Removed**: `markAsInput()`, `markAsOutput()`
- **Added**: `addTrackedRange()` (prompts for range name)

### 2. Model Structure
- `owner` and `description` fields removed
- `version` field required
- `monitoredCells` → `tracked_ranges`

### 3. Trace Structure
- No more generic "events" - only tracked range changes
- Must specify `tracked_range_name`
- No `type` field (input/output distinction removed)

## Versioning Behavior

### Creating a Model
```javascript
// First time
await upsertModel({
  model_name: "Revenue Model",
  tracked_ranges: [],
  model_id: "excel_123_abc"
});
// Result: {model_id: "excel_123_abc", version: 1}
```

### Updating a Model
```javascript
// Add tracked range
await upsertModel({
  model_name: "Revenue Model",
  tracked_ranges: [{name: "Revenue", range: "A1:A10"}],
  model_id: "excel_123_abc",
  version: 1  // Current version
});
// Result: {model_id: "excel_123_abc", version: 2}  ← Incremented!
```

## Rollback Plan

If you need to rollback to the old architecture:

1. **Keep old files**:
   - `domino-api.js` (do not delete)
   - `commands.js` (do not delete)
   - `domino-api-example.py` (old backend)

2. **Revert imports**:
   ```javascript
   // Change from:
   import { loadModel } from '../utils/domino-api-v2';

   // Back to:
   import { getModelById } from '../utils/domino-api';
   ```

3. **Start old backend**:
   ```bash
   python domino-api-example.py
   ```

## Testing Checklist

- [ ] Database tables created successfully
- [ ] New backend starts without errors
- [ ] Model registration creates model with version=1
- [ ] Workbook reload loads model successfully
- [ ] Adding tracked range increments version
- [ ] Changing tracked range creates trace
- [ ] Trace appears in backend logs
- [ ] Trace appears in UI "Recent Traces" section
- [ ] Removing tracked range updates model
- [ ] Offline queue works (test with backend stopped)

## Support & Troubleshooting

### Common Issues

**Issue**: "Model not found" error
- **Solution**: Model might not be registered. Click "Register Model" first.

**Issue**: Traces not appearing
- **Solution**: Make sure cell is in a tracked range. Use exact range matching.

**Issue**: Version not incrementing
- **Solution**: Make sure you're passing current `version` in update request.

**Issue**: API timeout errors
- **Solution**: Check `DOMINO_API_BASE` URL in both `domino-api-v2.js` and `commands-v2.js`.

### Debug Mode

Enable verbose logging:
```javascript
// In domino-api-v2.js
console.log('[domino-api-v2] Request:', data);
console.log('[domino-api-v2] Response:', result);

// In commands-v2.js
console.log('[commands-v2] Trace created:', traceData);
```

## Next Steps

1. **Production Database**:
   - Replace in-memory storage with SQL Server
   - Update connection string in `domino-api-backend.py`
   - Add connection pooling

2. **Authentication**:
   - Add JWT or OAuth to API
   - Protect endpoints with auth middleware

3. **Enhanced Tracking**:
   - Add formula tracking
   - Add worksheet-level tracking
   - Add calculation chain analysis

4. **UI Enhancements**:
   - Add bulk range import/export
   - Add visual range highlighting
   - Add trace filtering/search

## Questions?

Refer to `DEPLOYMENT.md` for the complete architecture specification.
